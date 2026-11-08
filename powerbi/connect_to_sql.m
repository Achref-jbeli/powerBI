// Power BI: Home > Get data > Blank query > Advanced Editor. One query per table (change the table name).
// Server/database must match load_excel_to_db.py.
let
    Source  = Sql.Database("localhost\SQLEXPRESS", "SalesKpiDW"),
    Table   = Source{[Schema = "dbo", Item = "FactSales"]}[Data]   // DimCustomer, DimProduct, DimDate, FactTarget
in
    Table

// Alternative (skip the database, read Excel directly) - same star schema can be built with Power Query:
// let
//     Source = Excel.Workbook(File.Contents("C:\...\data\sales_data.xlsx"), null, true),
//     Sales  = Table.PromoteHeaders(Source{[Item = "Sales", Kind = "Sheet"]}[Data]),
//     Typed  = Table.TransformColumnTypes(Sales, {{"OrderDate", type date}, {"Quantity", Int64.Type}, {"DiscountPct", type number}})
// in
//     Typed
