"""Creates test inbox folders: inbox/YYYY-MM-DD/sales_YYYY-MM-DD.xlsx (one per day).

Usage: python generate_daily_files.py [start_date YYYY-MM-DD] [days]
Example: python generate_daily_files.py 2026-11-01 7
"""
import random
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

start = date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else date(2026, 11, 1)
days = int(sys.argv[2]) if len(sys.argv) > 2 else 7
inbox = Path(__file__).parent / "inbox"

for n in range(days):
    day = start + timedelta(days=n)
    random.seed(day.toordinal())
    # OrderID = yymmdd * 1000 + n (fits in INT, unique per day, never collides with the history 1..3000)
    base = int(day.strftime("%y%m%d")) * 1000
    rows = [(base + i, day, random.randint(1, 40), random.randint(1, 20), random.randint(1, 10),
             random.choice([0, 0, 0, 0.05, 0.10, 0.15]))
            for i in range(random.randint(60, 140))]
    sales = pd.DataFrame(rows, columns=["OrderID", "OrderDate", "CustomerID", "ProductID", "Quantity", "DiscountPct"])
    folder = inbox / day.isoformat()
    folder.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(folder / f"sales_{day.isoformat()}.xlsx") as w:
        sales.to_excel(w, sheet_name="Sales", index=False)
    print("Created", folder.name, len(sales), "orders")
