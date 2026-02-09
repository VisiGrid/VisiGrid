-- VisiGrid Agent Demo: Multi-Sheet Workbook
-- Demonstrates writing to multiple sheets in a single workbook.
-- Run: vgrid sheet apply workbook.sheet --lua multi_sheet_model.lua --json
--
-- Multi-Sheet API (grid.* namespace):
--   grid.set{ sheet=N, cell="A1", value="..." }     -- N is 1-indexed
--   grid.set_batch{ sheet=N, cells={{cell="A1", value="..."}, ...} }
--   grid.format{ sheet=N, range="A1:B2", bold=true }
--
-- Sheets are auto-created as needed. If you reference sheet=3 and only
-- Sheet1 exists, Sheet2 and Sheet3 are automatically created.

--------------------------------------------------------------------------------
-- Sheet 1: Summary Dashboard
--------------------------------------------------------------------------------
set("A1", "Q1 Financial Summary")
meta("A1", { role = "title" })
style("A1", { bold = true })

set("A3", "Metric")
set("B3", "Value")
style("A3:B3", { bold = true })
meta("A3:B3", { role = "header" })

set("A4", "Total Revenue")
set("B4", "=SUM(Sheet2!B:B)")

set("A5", "Total Expenses")
set("B5", "=SUM(Sheet3!B:B)")

set("A6", "Net Income")
set("B6", "=B4-B5")
style("A6:B6", { bold = true })

--------------------------------------------------------------------------------
-- Sheet 2: Revenue Details
--------------------------------------------------------------------------------
grid.set{ sheet=2, cell="A1", value="Revenue by Category" }
grid.format{ sheet=2, range="A1", bold=true }

-- Headers
grid.set_batch{ sheet=2, cells={
    {cell="A3", value="Category"},
    {cell="B3", value="Amount"}
}}
grid.format{ sheet=2, range="A3:B3", bold=true }

-- Data rows
grid.set_batch{ sheet=2, cells={
    {cell="A4", value="Product Sales"},
    {cell="B4", value="150000"},
    {cell="A5", value="Services"},
    {cell="B5", value="75000"},
    {cell="A6", value="Subscriptions"},
    {cell="B6", value="45000"},
    {cell="A7", value="Licensing"},
    {cell="B7", value="30000"}
}}

-- Total
grid.set{ sheet=2, cell="A9", value="Total" }
grid.set{ sheet=2, cell="B9", value="=SUM(B4:B7)" }
grid.format{ sheet=2, range="A9:B9", bold=true }

--------------------------------------------------------------------------------
-- Sheet 3: Expense Details
--------------------------------------------------------------------------------
grid.set{ sheet=3, cell="A1", value="Expenses by Category" }
grid.format{ sheet=3, range="A1", bold=true }

-- Headers
grid.set_batch{ sheet=3, cells={
    {cell="A3", value="Category"},
    {cell="B3", value="Amount"}
}}
grid.format{ sheet=3, range="A3:B3", bold=true }

-- Data rows
grid.set_batch{ sheet=3, cells={
    {cell="A4", value="Salaries"},
    {cell="B4", value="120000"},
    {cell="A5", value="Infrastructure"},
    {cell="B5", value="35000"},
    {cell="A6", value="Marketing"},
    {cell="B6", value="25000"},
    {cell="A7", value="Operations"},
    {cell="B7", value="15000"}
}}

-- Total
grid.set{ sheet=3, cell="A9", value="Total" }
grid.set{ sheet=3, cell="B9", value="=SUM(B4:B7)" }
grid.format{ sheet=3, range="A9:B9", bold=true }
