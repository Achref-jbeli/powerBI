"""Daily loader.

Looks in the inbox folder for sub-folders named YYYY-MM-DD, each containing one .xlsx file
(sheets: Sales required; Customers, Products, Targets optional).

  - date already loaded successfully (LoadLog)  -> skipped, nothing is done
  - new date                                    -> loaded into the database (incremental)

Exit codes:  0 = at least one new date loaded   2 = nothing new   1 = a load failed

Usage: python load_daily.py [inbox_folder]
"""
import re
import sys
from pathlib import Path
from urllib.parse import quote_plus

import pandas as pd
from sqlalchemy import create_engine, text

CONN = (
    "DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost\\SQLEXPRESS;"
    "DATABASE=SalesKpiDW;Trusted_Connection=yes;TrustServerCertificate=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={quote_plus(CONN)}", fast_executemany=True)

inbox = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "inbox"
STAGING = {"Customers": "stg_Customers", "Products": "stg_Products", "Sales": "stg_Sales", "Targets": "stg_Targets"}
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def already_loaded(day: str) -> bool:
    with engine.connect() as cx:
        return cx.execute(text("SELECT 1 FROM LoadLog WHERE LoadDate = :d AND Status = 'Success'"), {"d": day}).first() is not None


def load_day(day: str, xlsx: Path):
    sheets = pd.read_excel(xlsx, sheet_name=None)
    if "Sales" not in sheets:
        raise ValueError("sheet 'Sales' is missing")
    with engine.begin() as cx:
        for sheet, table in STAGING.items():
            cx.execute(text(f"DELETE FROM {table}"))
            if sheet in sheets:
                df = sheets[sheet].dropna(how="all")
                df.columns = [str(c).strip() for c in df.columns]
                df.to_sql(table, cx, if_exists="append", index=False)
        cx.execute(text("EXEC usp_LoadDay :d, :f"), {"d": day, "f": xlsx.name})
        rows = cx.execute(text("SELECT TOP 1 RowsLoaded FROM LoadLog WHERE LoadDate=:d ORDER BY LogID DESC"), {"d": day}).scalar()
    return rows


def log_failure(day: str, name: str, msg: str):
    with engine.begin() as cx:
        cx.execute(text("INSERT LoadLog (LoadDate, FileName, RowsLoaded, Status, Message) VALUES (:d,:f,0,'Failed',:m)"),
                   {"d": day, "f": name, "m": msg[:900]})


loaded, failed = 0, 0
if not inbox.exists():
    print(f"Inbox folder not found: {inbox}")
    sys.exit(1)

for folder in sorted(p for p in inbox.iterdir() if p.is_dir() and DATE_RE.match(p.name)):
    day = folder.name
    if already_loaded(day):
        print(f"[{day}] already loaded -> skipped")
        continue
    files = sorted(folder.glob("*.xlsx"))
    if not files:
        print(f"[{day}] no .xlsx file in folder -> skipped")
        continue
    try:
        rows = load_day(day, files[0])
        print(f"[{day}] loaded {rows} new orders from {files[0].name}")
        loaded += 1
    except Exception as e:  # noqa: BLE001
        print(f"[{day}] FAILED: {e}")
        log_failure(day, files[0].name, str(e))
        failed += 1

print(f"Done. loaded={loaded} failed={failed}")
sys.exit(1 if failed else (0 if loaded else 2))
