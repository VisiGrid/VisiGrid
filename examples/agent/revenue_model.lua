-- VisiGrid Agent Demo: 12-Month Revenue Model
-- This script builds a complete revenue projection from scratch.
-- Run: vgrid sheet apply model.sheet --lua revenue_model.lua --json

-- Title
set("A1", "Revenue Projection Model")
meta("A1", { role = "title" })
style("A1", { bold = true })

-- Inputs section
set("A3", "Inputs")
meta("A3", { role = "section_header" })
style("A3", { bold = true })

set("A4", "Base Revenue")
set("B4", 100000)
meta("B4", { type = "input", unit = "usd" })

set("A5", "Monthly Growth Rate")
set("B5", 0.05)
meta("B5", { type = "input", unit = "percent" })

-- Monthly projections
set("A7", "Month")
set("B7", "Revenue")
set("C7", "Cumulative")
meta("A7:C7", { role = "header" })
style("A7:C7", { bold = true })

-- Month 1 (base)
set("A8", "Month 1")
set("B8", "=B4")
set("C8", "=B8")

-- Months 2-12 (growth formula)
set("A9", "Month 2")
set("B9", "=B8*(1+$B$5)")
set("C9", "=C8+B9")

set("A10", "Month 3")
set("B10", "=B9*(1+$B$5)")
set("C10", "=C9+B10")

set("A11", "Month 4")
set("B11", "=B10*(1+$B$5)")
set("C11", "=C10+B11")

set("A12", "Month 5")
set("B12", "=B11*(1+$B$5)")
set("C12", "=C11+B12")

set("A13", "Month 6")
set("B13", "=B12*(1+$B$5)")
set("C13", "=C12+B13")

set("A14", "Month 7")
set("B14", "=B13*(1+$B$5)")
set("C14", "=C13+B14")

set("A15", "Month 8")
set("B15", "=B14*(1+$B$5)")
set("C15", "=C14+B15")

set("A16", "Month 9")
set("B16", "=B15*(1+$B$5)")
set("C16", "=C15+B16")

set("A17", "Month 10")
set("B17", "=B16*(1+$B$5)")
set("C17", "=C16+B17")

set("A18", "Month 11")
set("B18", "=B17*(1+$B$5)")
set("C18", "=C17+B18")

set("A19", "Month 12")
set("B19", "=B18*(1+$B$5)")
set("C19", "=C18+B19")

-- Totals
set("A21", "Total")
set("B21", "=SUM(B8:B19)")
set("C21", "=C19")
meta("A21:C21", { role = "total" })
style("A21:C21", { bold = true })
