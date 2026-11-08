/* Step 2 - Database: staging tables + star schema + calculation views (SQL Server) */
IF DB_ID('SalesKpiDW') IS NULL CREATE DATABASE SalesKpiDW;
GO
USE SalesKpiDW;
GO

/* ---------- Staging (raw copy of Excel, reloaded on every run) ---------- */
DROP TABLE IF EXISTS stg_Sales, stg_Customers, stg_Products, stg_Targets;
CREATE TABLE stg_Customers (CustomerID INT, CustomerName NVARCHAR(100), City NVARCHAR(50), Segment NVARCHAR(30), Region NVARCHAR(30));
CREATE TABLE stg_Products  (ProductID INT, ProductName NVARCHAR(100), Category NVARCHAR(50), UnitPrice DECIMAL(12,2), UnitCost DECIMAL(12,2));
CREATE TABLE stg_Sales     (OrderID INT, OrderDate DATE, CustomerID INT, ProductID INT, Quantity INT, DiscountPct DECIMAL(5,2));
CREATE TABLE stg_Targets   (MonthStart DATE, RevenueTarget DECIMAL(14,2));
GO

/* ---------- Star schema ---------- */
DROP TABLE IF EXISTS FactSales, DimCustomer, DimProduct, DimDate, FactTarget;

CREATE TABLE DimCustomer (CustomerKey INT PRIMARY KEY, CustomerName NVARCHAR(100), City NVARCHAR(50), Region NVARCHAR(30), Segment NVARCHAR(30));
CREATE TABLE DimProduct  (ProductKey  INT PRIMARY KEY, ProductName  NVARCHAR(100), Category NVARCHAR(50), UnitPrice DECIMAL(12,2), UnitCost DECIMAL(12,2));
CREATE TABLE DimDate (
    DateKey INT PRIMARY KEY,            -- yyyymmdd
    [Date] DATE NOT NULL, [Year] INT, [Quarter] INT, [Month] INT,
    MonthName NVARCHAR(20), MonthStart DATE, YearMonth CHAR(7), WeekdayName NVARCHAR(20)
);
CREATE TABLE FactSales (
    OrderID INT PRIMARY KEY,
    DateKey INT NOT NULL REFERENCES DimDate(DateKey),
    CustomerKey INT NOT NULL REFERENCES DimCustomer(CustomerKey),
    ProductKey INT NOT NULL REFERENCES DimProduct(ProductKey),
    Quantity INT, DiscountPct DECIMAL(5,2),
    GrossAmount DECIMAL(14,2),          -- Qty * UnitPrice
    NetAmount   DECIMAL(14,2),          -- Gross * (1 - discount)
    CostAmount  DECIMAL(14,2),          -- Qty * UnitCost
    Profit      AS (NetAmount - CostAmount) PERSISTED
);
CREATE TABLE FactTarget (MonthStart DATE PRIMARY KEY, RevenueTarget DECIMAL(14,2));
GO

/* ---------- Transform: staging -> star schema (run after each load) ---------- */
CREATE OR ALTER PROCEDURE usp_BuildWarehouse AS
BEGIN
    SET NOCOUNT ON;
    DELETE FROM FactSales; DELETE FROM FactTarget; DELETE FROM DimCustomer; DELETE FROM DimProduct; DELETE FROM DimDate;

    INSERT DimCustomer SELECT DISTINCT CustomerID, CustomerName, City, Region, Segment FROM stg_Customers;
    INSERT DimProduct  SELECT DISTINCT ProductID, ProductName, Category, UnitPrice, UnitCost FROM stg_Products;

    -- Calendar covering the full data range
    DECLARE @min DATE = (SELECT DATEFROMPARTS(YEAR(MIN(OrderDate)),1,1) FROM stg_Sales),
            @max DATE = (SELECT DATEFROMPARTS(YEAR(MAX(OrderDate)),12,31) FROM stg_Sales);
    ;WITH d AS (SELECT @min AS dt UNION ALL SELECT DATEADD(DAY,1,dt) FROM d WHERE dt < @max)
    INSERT DimDate
    SELECT CONVERT(INT, FORMAT(dt,'yyyyMMdd')), dt, YEAR(dt), DATEPART(QUARTER,dt), MONTH(dt),
           DATENAME(MONTH,dt), DATEFROMPARTS(YEAR(dt),MONTH(dt),1), FORMAT(dt,'yyyy-MM'), DATENAME(WEEKDAY,dt)
    FROM d OPTION (MAXRECURSION 0);

    -- Data-quality rules: drop orphans, non-positive quantities, duplicate orders
    INSERT FactSales (OrderID, DateKey, CustomerKey, ProductKey, Quantity, DiscountPct, GrossAmount, NetAmount, CostAmount)
    SELECT s.OrderID, CONVERT(INT, FORMAT(s.OrderDate,'yyyyMMdd')), s.CustomerID, s.ProductID, s.Quantity, s.DiscountPct,
           s.Quantity * p.UnitPrice,
           s.Quantity * p.UnitPrice * (1 - s.DiscountPct),
           s.Quantity * p.UnitCost
    FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY OrderDate) rn FROM stg_Sales) s
    JOIN DimProduct p  ON p.ProductKey  = s.ProductID
    JOIN DimCustomer c ON c.CustomerKey = s.CustomerID
    WHERE s.rn = 1 AND s.Quantity > 0;

    INSERT FactTarget SELECT MonthStart, RevenueTarget FROM stg_Targets;
END
GO

/* ---------- Calculation layer (SQL side KPIs) ---------- */
CREATE OR ALTER VIEW vw_MonthlyKpi AS
SELECT d.MonthStart, d.YearMonth,
       SUM(f.NetAmount)  AS Revenue,
       SUM(f.Profit)     AS Profit,
       COUNT(*)          AS Orders,
       SUM(f.Quantity)   AS Units,
       SUM(f.Profit) / NULLIF(SUM(f.NetAmount),0) AS MarginPct,
       SUM(f.NetAmount) / NULLIF(COUNT(*),0)      AS AvgOrderValue,
       MAX(t.RevenueTarget)                       AS RevenueTarget,
       SUM(f.NetAmount) / NULLIF(MAX(t.RevenueTarget),0) AS TargetAttainment,
       LAG(SUM(f.NetAmount)) OVER (ORDER BY d.MonthStart) AS PrevMonthRevenue,
       SUM(f.NetAmount) / NULLIF(LAG(SUM(f.NetAmount)) OVER (ORDER BY d.MonthStart),0) - 1 AS MoMGrowth
FROM FactSales f
JOIN DimDate d ON d.DateKey = f.DateKey
LEFT JOIN FactTarget t ON t.MonthStart = d.MonthStart
GROUP BY d.MonthStart, d.YearMonth;
GO

CREATE OR ALTER VIEW vw_CustomerRfm AS      -- RFM scoring (1-5 each)
WITH base AS (
    SELECT c.CustomerKey, c.CustomerName, c.Region,
           DATEDIFF(DAY, MAX(d.[Date]), (SELECT MAX([Date]) FROM DimDate d2 JOIN FactSales f2 ON f2.DateKey=d2.DateKey)) AS RecencyDays,
           COUNT(*) AS Frequency, SUM(f.NetAmount) AS Monetary
    FROM FactSales f JOIN DimCustomer c ON c.CustomerKey=f.CustomerKey JOIN DimDate d ON d.DateKey=f.DateKey
    GROUP BY c.CustomerKey, c.CustomerName, c.Region)
SELECT *, 6 - NTILE(5) OVER (ORDER BY RecencyDays) AS R,
          NTILE(5) OVER (ORDER BY Frequency) AS F,
          NTILE(5) OVER (ORDER BY Monetary)  AS M
FROM base;
GO
