-- Bond Portfolio: Accrued Interest Demo
-- Demonstrates custom functions in a verifiable spreadsheet.
--
-- Prerequisites:
--   cp examples/custom-functions/functions.lua ~/.config/visigrid/functions.lua
--
-- Build:
--   visigrid-cli sheet apply portfolio.sheet --lua bond_portfolio.lua --stamp --json
--
-- The formulas =ACCRUED_INTEREST(...) and =HAIRCUT(...) evaluate in the GUI
-- when custom functions are loaded. The fingerprint covers inputs + formulas,
-- proving that identical inputs always produce identical computation.

-- Title
set("A1", "Bond Portfolio \u{2014} Accrued Interest")
meta("A1", { role = "title" })
style("A1", { bold = true })

-- Inputs
set("A3", "Settlement")
set("B3", "2025-06-15")
meta("B3", { type = "input" })

-- Column headers
set("A5", "Bond")
set("B5", "Principal")
set("C5", "Rate")
set("D5", "Days")
set("E5", "Accrued")
set("F5", "Haircut %")
set("G5", "Net Value")
meta("A5:G5", { role = "header" })
style("A5:G5", { bold = true })

-- Bond 1: US Treasury 10Y
set("A6", "UST 10Y")
set("B6", 1000000)
set("C6", 0.0425)
set("D6", 92)
set("E6", "=ACCRUED_INTEREST(B6, C6, D6)")
set("F6", 0.02)
set("G6", "=HAIRCUT(B6 + E6, F6)")
meta("B6", { type = "input", unit = "usd" })
meta("C6", { type = "input", unit = "rate" })
meta("D6", { type = "input", unit = "days" })

-- Bond 2: Corp Bond AA
set("A7", "Corp AA 5Y")
set("B7", 500000)
set("C7", 0.055)
set("D7", 45)
set("E7", "=ACCRUED_INTEREST(B7, C7, D7)")
set("F7", 0.05)
set("G7", "=HAIRCUT(B7 + E7, F7)")
meta("B7", { type = "input", unit = "usd" })
meta("C7", { type = "input", unit = "rate" })
meta("D7", { type = "input", unit = "days" })

-- Bond 3: Muni Bond
set("A8", "Muni GO 20Y")
set("B8", 750000)
set("C8", 0.035)
set("D8", 120)
set("E8", "=ACCRUED_INTEREST(B8, C8, D8)")
set("F8", 0.03)
set("G8", "=HAIRCUT(B8 + E8, F8)")
meta("B8", { type = "input", unit = "usd" })
meta("C8", { type = "input", unit = "rate" })
meta("D8", { type = "input", unit = "days" })

-- Totals
set("A10", "Total")
set("B10", "=SUM(B6:B8)")
set("E10", "=SUM(E6:E8)")
set("G10", "=SUM(G6:G8)")
meta("A10:G10", { role = "total" })
style("A10:G10", { bold = true })
