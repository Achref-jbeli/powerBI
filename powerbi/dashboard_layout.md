# Step 5 - Dashboard layout (Power BI Desktop, 16:9 canvas)

## Page 1 - Executive Overview
```
+-----------------------------------------------------------------------+
| Title (dynamic: [Title Subtitle])        Slicers: Year | Region | Segment |
+----------+----------+----------+----------+----------+----------------+
| Revenue  | Profit   | Margin % | Orders   | AOV      | Target Attain. |  <- KPI cards (+ YoY / MoM small text)
+----------+----------+----------+----------+----------+----------------+
| Line+column: Revenue (cols) vs Revenue Target (line) by DimDate[YearMonth] |
+-----------------------------------+-----------------------------------+
| Bar: Revenue by Category          | Map / filled map: Revenue by City  |
+-----------------------------------+-----------------------------------+
| Table: Top 10 products (Revenue, Margin %, Rank) with data bars         |
+-----------------------------------------------------------------------+
```

## Page 2 - Customers
- Matrix of `vw_CustomerRfm` (R, F, M) with heat-map conditional formatting
- Scatter: Frequency (x) vs Monetary (y), size = Orders, legend = Segment
- Card: Customers, New Customers; Donut: Revenue by Segment

## Page 3 - Target tracking
- Gauge: `[Target Attainment %]` (max 1.2, target 1)
- Column chart by month with colour from `[Target Color]` (Format > Columns > fx > Field value)
- Table: month, Revenue, Target, Gap, Status

## Interactivity checklist
- Sync the Year/Region slicers across pages (View > Sync slicers)
- Drill-through page: right-click a product -> Product detail
- Bookmarks for "Last 12 months" / "Full period"
- Tooltip page showing a monthly sparkline

## Refresh
- Publish to Power BI Service, install an On-premises data gateway for the local SQL Server,
  schedule refresh after `python load_excel_to_db.py` (e.g. via Windows Task Scheduler).

## Page 4 - Daily trend (daily-load flow)
- Cards: `[Revenue Last Day]`, `[Orders Last Day]`, `[Trend Arrow]`, `[Last Load Date]`, `[Last Refresh Run]`
- Line chart: X = `DimDate[Date]` (continuous), Y = `[Revenue]` + `[Revenue 7D Avg]`;
  visual filter **Show Date = 1** (otherwise the axis runs to 31 Dec)
- Column chart: `[DoD Growth %]` per day (conditional colour: green >= 0, red < 0)
- Slicer: DimDate[Date] "Relative date" -> last 30 days
- Table: LoadLog (LoadDate, FileName, RowsLoaded, Status) to monitor the loads
