-- VisiGrid Agent Demo: Reusable Template for `vgrid fill`
--
-- This script builds a template .sheet with headers, a totals row
-- with formulas, and styling. The data area (rows 2–N) is left empty
-- for `vgrid fill` to populate with CSV data.
--
-- Two-step workflow:
--
--   Step 1: Build the template
--   vgrid sheet apply template.sheet --lua fill_template.lua --json
--
--   Step 2: Fill with data (headers excluded, data injected at A2)
--   vgrid fill template.sheet --csv data.csv --target A2 --headers --out filled.sheet --json
--
-- The template expects CSV columns: Date, Description, Category, Amount, Status
-- Totals row is placed at row 102 (supports up to 100 data rows).
-- Adjust the totals row if your dataset is larger.
--
-- Run: vgrid sheet apply template.sheet --lua fill_template.lua --json

-- ── Title ──────────────────────────────────────────────────────

-- (Title goes above the data table so fill --target A3 can skip it)
-- If you want a title row, put headers at row 2 and data at row 3:
--   vgrid fill template.sheet --csv data.csv --target A3 --headers --out filled.sheet --json

-- ── Header row ─────────────────────────────────────────────────

set("A1", "Date")
set("B1", "Description")
set("C1", "Category")
set("D1", "Amount")
set("E1", "Status")

meta("A1:E1", { role = "header" })
style("A1:E1", { bold = true })

-- ── Totals / summary row ───────────────────────────────────────
-- Placed at row 102 to accommodate up to 100 data rows (A2:A101).
-- Formulas reference the full data range; empty cells are ignored by
-- SUM, AVERAGE, COUNTA, etc.

local totals_row = 102

set("A" .. totals_row, "TOTALS")
style("A" .. totals_row, { bold = true })

-- Total amount
set("D" .. totals_row, "=SUM(D2:D101)")
meta("D" .. totals_row, { role = "total", formula = "sum_amount" })
style("D" .. totals_row, { bold = true })

-- Record count (non-empty dates = number of filled rows)
set("E" .. totals_row, "=COUNTA(A2:A101)")
meta("E" .. totals_row, { role = "total", formula = "record_count" })

-- ── Summary statistics row ─────────────────────────────────────

local stats_row = 103

set("A" .. stats_row, "STATS")
style("A" .. stats_row, { bold = true })

-- Average amount
set("C" .. stats_row, "Average")
set("D" .. stats_row, "=AVERAGE(D2:D101)")
meta("D" .. stats_row, { role = "summary", formula = "avg_amount" })

-- ── Metadata ───────────────────────────────────────────────────
-- Tag the template so agents can identify it programmatically.

meta("A1", { template = "fill_demo", version = "1", columns = 5, data_start = "A2", data_end = "E101" })
