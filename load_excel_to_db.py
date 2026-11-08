"""Step 3 - ETL: read every sheet of the Excel file and load it into SQL Server staging,
then call usp_BuildWarehouse to build the star schema.

Usage:  python load_excel_to_db.py [path/to/file.xlsx]
Edit CONN below for your server (default: local SQL Server / SQLEXPRESS, Windows auth).
"""
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

xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "data" / "sales_data.xlsx"
sheets = pd.read_excel(xlsx, sheet_name=None)          # {sheet name: DataFrame}

mapping = {"Customers": "stg_Customers", "Products": "stg_Products", "Sales": "stg_Sales", "Targets": "stg_Targets"}

with engine.begin() as cx:
    for sheet, table in mapping.items():
        df = sheets[sheet].dropna(how="all")
        df.columns = [c.strip() for c in df.columns]
        cx.execute(text(f"DELETE FROM {table}"))
        df.to_sql(table, cx, if_exists="append", index=False)
        print(f"{sheet:10} -> {table:14} {len(df):>6} rows")
    cx.execute(text("EXEC usp_BuildWarehouse"))
    print("Warehouse rebuilt.")
    print(pd.read_sql("SELECT TOP 5 * FROM vw_MonthlyKpi ORDER BY MonthStart", cx))
