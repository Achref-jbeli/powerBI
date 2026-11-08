"""Step 1 - Generate a sample Excel workbook (the 'messy business source').

Creates data/sales_data.xlsx with 3 sheets: Customers, Products, Sales.
Replace this file with your real Excel workbook later; keep the column names
(or adjust load_excel_to_db.py).
"""
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)
out = Path(__file__).parent / "data"
out.mkdir(exist_ok=True)

regions = {"Tunis": "North", "Sfax": "South", "Sousse": "Center", "Bizerte": "North", "Gabes": "South"}
customers = pd.DataFrame(
    [(i, f"Customer {i:03d}", random.choice(list(regions)), random.choice(["Retail", "Wholesale", "Online"]))
     for i in range(1, 41)],
    columns=["CustomerID", "CustomerName", "City", "Segment"],
)
customers["Region"] = customers["City"].map(regions)

cats = {"Laptops": (900, 1800), "Phones": (300, 1200), "Accessories": (10, 120), "Monitors": (150, 600)}
products = []
pid = 1
for cat, (lo, hi) in cats.items():
    for n in range(1, 6):
        price = round(random.uniform(lo, hi), 2)
        products.append((pid, f"{cat[:-1]} {n}", cat, price, round(price * random.uniform(0.55, 0.8), 2)))
        pid += 1
products = pd.DataFrame(products, columns=["ProductID", "ProductName", "Category", "UnitPrice", "UnitCost"])

rows, start = [], date(2025, 1, 1)
for oid in range(1, 3001):
    p = products.sample(1).iloc[0]
    qty = random.randint(1, 10)
    rows.append((
        oid,
        start + timedelta(days=random.randint(0, 600)),
        random.randint(1, 40),
        int(p.ProductID),
        qty,
        random.choice([0, 0, 0, 0.05, 0.10, 0.15]),
    ))
sales = pd.DataFrame(rows, columns=["OrderID", "OrderDate", "CustomerID", "ProductID", "Quantity", "DiscountPct"])

with pd.ExcelWriter(out / "sales_data.xlsx") as w:
    customers.to_excel(w, sheet_name="Customers", index=False)
    products.to_excel(w, sheet_name="Products", index=False)
    sales.to_excel(w, sheet_name="Sales", index=False)

# Monthly revenue targets for the "actual vs target" KPI
months = pd.period_range("2025-01", "2026-08", freq="M")
targets = pd.DataFrame({"MonthStart": months.to_timestamp(), "RevenueTarget": [random.randint(180, 260) * 1000 for _ in months]})
with pd.ExcelWriter(out / "sales_data.xlsx", mode="a") as w:
    targets.to_excel(w, sheet_name="Targets", index=False)

print("Created", out / "sales_data.xlsx")
