/* Step 2b - Daily incremental loading (safe to re-run, does NOT wipe data) */
USE SalesKpiDW;
GO
SET QUOTED_IDENTIFIER ON;   -- required by the persisted computed column FactSales.Profit (sqlcmd defaults to OFF)
SET ANSI_NULLS ON;
GO

IF OBJECT_ID('dbo.LoadLog') IS NULL
CREATE TABLE dbo.LoadLog (
    LogID      INT IDENTITY PRIMARY KEY,
    LoadDate   DATE NOT NULL,                 -- date taken from the folder name
    FileName   NVARCHAR(260),
    RowsLoaded INT,
    Status     NVARCHAR(10) NOT NULL,         -- Success | Failed
    Message    NVARCHAR(1000) NULL,
    LoadedAt   DATETIME2 NOT NULL DEFAULT SYSDATETIME()
);
GO

/* Called once per date folder, after the Excel sheets were copied into stg_* tables. */
CREATE OR ALTER PROCEDURE dbo.usp_LoadDay @LoadDate DATE, @FileName NVARCHAR(260) AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRAN;

    -- 1) Dimensions: insert new, update changed
    MERGE DimCustomer t USING (SELECT DISTINCT CustomerID, CustomerName, City, Region, Segment FROM stg_Customers) s
      ON t.CustomerKey = s.CustomerID
    WHEN MATCHED THEN UPDATE SET CustomerName=s.CustomerName, City=s.City, Region=s.Region, Segment=s.Segment
    WHEN NOT MATCHED THEN INSERT (CustomerKey, CustomerName, City, Region, Segment) VALUES (s.CustomerID, s.CustomerName, s.City, s.Region, s.Segment);

    MERGE DimProduct t USING (SELECT DISTINCT ProductID, ProductName, Category, UnitPrice, UnitCost FROM stg_Products) s
      ON t.ProductKey = s.ProductID
    WHEN MATCHED THEN UPDATE SET ProductName=s.ProductName, Category=s.Category, UnitPrice=s.UnitPrice, UnitCost=s.UnitCost
    WHEN NOT MATCHED THEN INSERT (ProductKey, ProductName, Category, UnitPrice, UnitCost) VALUES (s.ProductID, s.ProductName, s.Category, s.UnitPrice, s.UnitCost);

    -- 2) Calendar: make sure every year touched by the data exists
    DECLARE @min DATE = (SELECT DATEFROMPARTS(YEAR(MIN(OrderDate)),1,1) FROM stg_Sales),
            @max DATE = (SELECT DATEFROMPARTS(YEAR(MAX(OrderDate)),12,31) FROM stg_Sales);
    IF @min IS NOT NULL
    BEGIN
    ;WITH d AS (SELECT @min AS dt UNION ALL SELECT DATEADD(DAY,1,dt) FROM d WHERE dt < @max)
    INSERT DimDate
    SELECT CONVERT(INT, FORMAT(dt,'yyyyMMdd')), dt, YEAR(dt), DATEPART(QUARTER,dt), MONTH(dt),
           DATENAME(MONTH,dt), DATEFROMPARTS(YEAR(dt),MONTH(dt),1), FORMAT(dt,'yyyy-MM'), DATENAME(WEEKDAY,dt)
    FROM d
    WHERE NOT EXISTS (SELECT 1 FROM DimDate x WHERE x.[Date] = d.dt)
    OPTION (MAXRECURSION 0);
    END

    -- 3) Facts: append only orders that do not exist yet (re-sending an order is harmless)
    INSERT FactSales (OrderID, DateKey, CustomerKey, ProductKey, Quantity, DiscountPct, GrossAmount, NetAmount, CostAmount)
    SELECT s.OrderID, CONVERT(INT, FORMAT(s.OrderDate,'yyyyMMdd')), s.CustomerID, s.ProductID, s.Quantity, s.DiscountPct,
           s.Quantity * p.UnitPrice, s.Quantity * p.UnitPrice * (1 - s.DiscountPct), s.Quantity * p.UnitCost
    FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY OrderDate) rn FROM stg_Sales) s
    JOIN DimProduct p  ON p.ProductKey  = s.ProductID
    JOIN DimCustomer c ON c.CustomerKey = s.CustomerID
    WHERE s.rn = 1 AND s.Quantity > 0
      AND NOT EXISTS (SELECT 1 FROM FactSales f WHERE f.OrderID = s.OrderID);
    DECLARE @rows INT = @@ROWCOUNT;

    -- 4) Targets: insert / update
    MERGE FactTarget t USING (SELECT DISTINCT MonthStart, RevenueTarget FROM stg_Targets) s ON t.MonthStart = s.MonthStart
    WHEN MATCHED THEN UPDATE SET RevenueTarget = s.RevenueTarget
    WHEN NOT MATCHED THEN INSERT (MonthStart, RevenueTarget) VALUES (s.MonthStart, s.RevenueTarget);

    INSERT LoadLog (LoadDate, FileName, RowsLoaded, Status) VALUES (@LoadDate, @FileName, @rows, 'Success');
    COMMIT;
END
GO

/* Daily trend with day-over-day change and 7-day moving average */
CREATE OR ALTER VIEW dbo.vw_DailyKpi AS
WITH daily AS (
    SELECT d.[Date], SUM(f.NetAmount) AS Revenue, SUM(f.Profit) AS Profit,
           COUNT(*) AS Orders, SUM(f.Quantity) AS Units
    FROM FactSales f JOIN DimDate d ON d.DateKey = f.DateKey
    GROUP BY d.[Date])
SELECT [Date], Revenue, Profit, Orders, Units,
       Profit / NULLIF(Revenue,0) AS MarginPct,
       Revenue / NULLIF(LAG(Revenue) OVER (ORDER BY [Date]),0) - 1 AS DoDGrowth,
       AVG(Revenue) OVER (ORDER BY [Date] ROWS BETWEEN 6 PRECEDING AND CURRENT ROW) AS Revenue7DAvg
FROM daily;
GO
