-- Microsoft 3-Statement Financial Model
-- FY2022-FY2024: SEC 10-K sourced actuals
-- FY2025-FY2029: Formula-driven projections
-- Run: visigrid-cli sheet apply msft.sheet --lua msft_3statement_model.lua --json

--------------------------------------------------------------------------------
-- Name the sheets
--------------------------------------------------------------------------------
grid.name_sheet{ sheet=1, name="Assumptions" }
grid.name_sheet{ sheet=2, name="Income Statement" }
grid.name_sheet{ sheet=3, name="Balance Sheet" }
grid.name_sheet{ sheet=4, name="Cash Flow" }

--------------------------------------------------------------------------------
-- Sheet 1: Assumptions & Drivers
--------------------------------------------------------------------------------
set("A1", "Microsoft Corporation - 3-Statement Model")
meta("A1", { role = "title" })
style("A1", { bold = true })

set("A3", "Model Assumptions")
meta("A3", { role = "section_header" })
style("A3", { bold = true })

-- Column headers
set("B5", "FY2025")
set("C5", "FY2026")
set("D5", "FY2027")
set("E5", "FY2028")
set("F5", "FY2029")
style("B5:F5", { bold = true })
meta("B5:F5", { role = "header" })

-- Revenue growth rate
set("A6", "Revenue Growth %")
meta("A6", { type = "driver" })
set("B6", 0.12)
set("C6", 0.10)
set("D6", 0.09)
set("E6", 0.08)
set("F6", 0.07)
meta("B6:F6", { type = "input", unit = "percent" })

-- Operating margin
set("A7", "Operating Margin %")
meta("A7", { type = "driver" })
set("B7", 0.44)
set("C7", 0.445)
set("D7", 0.45)
set("E7", 0.455)
set("F7", 0.46)
meta("B7:F7", { type = "input", unit = "percent" })

-- Tax rate
set("A8", "Effective Tax Rate %")
meta("A8", { type = "driver" })
set("B8", 0.18)
set("C8", 0.18)
set("D8", 0.18)
set("E8", 0.18)
set("F8", 0.18)
meta("B8:F8", { type = "input", unit = "percent" })

-- Interest rate on debt
set("A9", "Interest Rate on Debt %")
meta("A9", { type = "driver" })
set("B9", 0.035)
set("C9", 0.035)
set("D9", 0.035)
set("E9", 0.035)
set("F9", 0.035)
meta("B9:F9", { type = "input", unit = "percent" })

-- D&A as % of revenue
set("A10", "D&A % of Revenue")
meta("A10", { type = "driver" })
set("B10", 0.07)
set("C10", 0.07)
set("D10", 0.07)
set("E10", 0.07)
set("F10", 0.07)
meta("B10:F10", { type = "input", unit = "percent" })

-- CapEx as % of revenue
set("A11", "CapEx % of Revenue")
meta("A11", { type = "driver" })
set("B11", 0.12)
set("C11", 0.12)
set("D11", 0.11)
set("E11", 0.11)
set("F11", 0.10)
meta("B11:F11", { type = "input", unit = "percent" })

-- NWC as % of revenue
set("A12", "NWC % of Revenue")
meta("A12", { type = "driver" })
set("B12", -0.15)
set("C12", -0.15)
set("D12", -0.15)
set("E12", -0.15)
set("F12", -0.15)
meta("B12:F12", { type = "input", unit = "percent" })

-- Dividend payout
set("A13", "Dividend Payout ($M)")
meta("A13", { type = "driver" })
set("B13", 22000)
set("C13", 23000)
set("D13", 24000)
set("E13", 25000)
set("F13", 26000)
meta("B13:F13", { type = "input", unit = "usd_millions" })

-- Share buybacks
set("A14", "Share Buybacks ($M)")
meta("A14", { type = "driver" })
set("B14", 20000)
set("C14", 22000)
set("D14", 24000)
set("E14", 26000)
set("F14", 28000)
meta("B14:F14", { type = "input", unit = "usd_millions" })

--------------------------------------------------------------------------------
-- Sheet 2: Income Statement
--------------------------------------------------------------------------------
grid.set{ sheet=2, cell="A1", value="Income Statement" }
grid.format{ sheet=2, range="A1", bold=true }

-- Year headers (historical + projected)
grid.set_batch{ sheet=2, cells={
    {cell="A3", value="($ in millions)"},
    {cell="B3", value="FY2022"},
    {cell="C3", value="FY2023"},
    {cell="D3", value="FY2024"},
    {cell="E3", value="FY2025"},
    {cell="F3", value="FY2026"},
    {cell="G3", value="FY2027"},
    {cell="H3", value="FY2028"},
    {cell="I3", value="FY2029"}
}}
grid.format{ sheet=2, range="A3:I3", bold=true }

-- Historical marker
grid.set{ sheet=2, cell="B4", value="Actual" }
grid.set{ sheet=2, cell="C4", value="Actual" }
grid.set{ sheet=2, cell="D4", value="Actual" }
grid.set{ sheet=2, cell="E4", value="Projected" }
grid.set{ sheet=2, cell="F4", value="Projected" }
grid.set{ sheet=2, cell="G4", value="Projected" }
grid.set{ sheet=2, cell="H4", value="Projected" }
grid.set{ sheet=2, cell="I4", value="Projected" }

-- Revenue (SEC 10-K actuals for FY22-24, projected for FY25-29)
grid.set{ sheet=2, cell="A6", value="Revenue" }
grid.set{ sheet=2, cell="B6", value=198270 }  -- FY2022 actual
grid.set{ sheet=2, cell="C6", value=211915 }  -- FY2023 actual
grid.set{ sheet=2, cell="D6", value=245122 }  -- FY2024 actual
grid.set{ sheet=2, cell="E6", value="=D6*(1+Assumptions!B6)" }  -- FY2025 projected
grid.set{ sheet=2, cell="F6", value="=E6*(1+Assumptions!C6)" }  -- FY2026 projected
grid.set{ sheet=2, cell="G6", value="=F6*(1+Assumptions!D6)" }  -- FY2027 projected
grid.set{ sheet=2, cell="H6", value="=G6*(1+Assumptions!E6)" }  -- FY2028 projected
grid.set{ sheet=2, cell="I6", value="=H6*(1+Assumptions!F6)" }  -- FY2029 projected
grid.format{ sheet=2, range="A6", bold=true }

-- Cost of Revenue
grid.set{ sheet=2, cell="A7", value="Cost of Revenue" }
grid.set{ sheet=2, cell="B7", value=-62650 }  -- FY2022
grid.set{ sheet=2, cell="C7", value=-65863 }  -- FY2023
grid.set{ sheet=2, cell="D7", value=-74114 }  -- FY2024
grid.set{ sheet=2, cell="E7", value="=-E6*(1-Assumptions!B7)*(D7/-D6)" }
grid.set{ sheet=2, cell="F7", value="=-F6*(1-Assumptions!C7)*(D7/-D6)" }
grid.set{ sheet=2, cell="G7", value="=-G6*(1-Assumptions!D7)*(D7/-D6)" }
grid.set{ sheet=2, cell="H7", value="=-H6*(1-Assumptions!E7)*(D7/-D6)" }
grid.set{ sheet=2, cell="I7", value="=-I6*(1-Assumptions!F7)*(D7/-D6)" }

-- Gross Profit
grid.set{ sheet=2, cell="A8", value="Gross Profit" }
grid.set{ sheet=2, cell="B8", value="=B6+B7" }
grid.set{ sheet=2, cell="C8", value="=C6+C7" }
grid.set{ sheet=2, cell="D8", value="=D6+D7" }
grid.set{ sheet=2, cell="E8", value="=E6+E7" }
grid.set{ sheet=2, cell="F8", value="=F6+F7" }
grid.set{ sheet=2, cell="G8", value="=G6+G7" }
grid.set{ sheet=2, cell="H8", value="=H6+H7" }
grid.set{ sheet=2, cell="I8", value="=I6+I7" }
grid.format{ sheet=2, range="A8:I8", bold=true }

-- Operating Expenses
grid.set{ sheet=2, cell="A10", value="Research & Development" }
grid.set{ sheet=2, cell="B10", value=-24512 }
grid.set{ sheet=2, cell="C10", value=-27195 }
grid.set{ sheet=2, cell="D10", value=-29510 }
grid.set{ sheet=2, cell="E10", value="=E6*(D10/D6)" }
grid.set{ sheet=2, cell="F10", value="=F6*(D10/D6)" }
grid.set{ sheet=2, cell="G10", value="=G6*(D10/D6)" }
grid.set{ sheet=2, cell="H10", value="=H6*(D10/D6)" }
grid.set{ sheet=2, cell="I10", value="=I6*(D10/D6)" }

grid.set{ sheet=2, cell="A11", value="Sales & Marketing" }
grid.set{ sheet=2, cell="B11", value=-21825 }
grid.set{ sheet=2, cell="C11", value=-22759 }
grid.set{ sheet=2, cell="D11", value=-24456 }
grid.set{ sheet=2, cell="E11", value="=E6*(D11/D6)" }
grid.set{ sheet=2, cell="F11", value="=F6*(D11/D6)" }
grid.set{ sheet=2, cell="G11", value="=G6*(D11/D6)" }
grid.set{ sheet=2, cell="H11", value="=H6*(D11/D6)" }
grid.set{ sheet=2, cell="I11", value="=I6*(D11/D6)" }

grid.set{ sheet=2, cell="A12", value="General & Administrative" }
grid.set{ sheet=2, cell="B12", value=-5900 }
grid.set{ sheet=2, cell="C12", value=-7575 }
grid.set{ sheet=2, cell="D12", value=-7609 }
grid.set{ sheet=2, cell="E12", value="=E6*(D12/D6)" }
grid.set{ sheet=2, cell="F12", value="=F6*(D12/D6)" }
grid.set{ sheet=2, cell="G12", value="=G6*(D12/D6)" }
grid.set{ sheet=2, cell="H12", value="=H6*(D12/D6)" }
grid.set{ sheet=2, cell="I12", value="=I6*(D12/D6)" }

-- Total Operating Expenses
grid.set{ sheet=2, cell="A13", value="Total Operating Expenses" }
grid.set{ sheet=2, cell="B13", value="=B10+B11+B12" }
grid.set{ sheet=2, cell="C13", value="=C10+C11+C12" }
grid.set{ sheet=2, cell="D13", value="=D10+D11+D12" }
grid.set{ sheet=2, cell="E13", value="=E10+E11+E12" }
grid.set{ sheet=2, cell="F13", value="=F10+F11+F12" }
grid.set{ sheet=2, cell="G13", value="=G10+G11+G12" }
grid.set{ sheet=2, cell="H13", value="=H10+H11+H12" }
grid.set{ sheet=2, cell="I13", value="=I10+I11+I12" }
grid.format{ sheet=2, range="A13:I13", bold=true }

-- Operating Income
grid.set{ sheet=2, cell="A15", value="Operating Income" }
grid.set{ sheet=2, cell="B15", value="=B8+B13" }
grid.set{ sheet=2, cell="C15", value="=C8+C13" }
grid.set{ sheet=2, cell="D15", value="=D8+D13" }
-- Projected: use operating margin assumption
grid.set{ sheet=2, cell="E15", value="=E6*Assumptions!B7" }
grid.set{ sheet=2, cell="F15", value="=F6*Assumptions!C7" }
grid.set{ sheet=2, cell="G15", value="=G6*Assumptions!D7" }
grid.set{ sheet=2, cell="H15", value="=H6*Assumptions!E7" }
grid.set{ sheet=2, cell="I15", value="=I6*Assumptions!F7" }
grid.format{ sheet=2, range="A15:I15", bold=true }

-- Interest Expense (driven by prior year debt * rate)
grid.set{ sheet=2, cell="A17", value="Interest Expense" }
grid.set{ sheet=2, cell="B17", value=-2063 }
grid.set{ sheet=2, cell="C17", value=-1968 }
grid.set{ sheet=2, cell="D17", value=-1602 }
grid.set{ sheet=2, cell="E17", value="=-'Balance Sheet'!D20*Assumptions!B9" }
grid.set{ sheet=2, cell="F17", value="=-'Balance Sheet'!E20*Assumptions!C9" }
grid.set{ sheet=2, cell="G17", value="=-'Balance Sheet'!F20*Assumptions!D9" }
grid.set{ sheet=2, cell="H17", value="=-'Balance Sheet'!G20*Assumptions!E9" }
grid.set{ sheet=2, cell="I17", value="=-'Balance Sheet'!H20*Assumptions!F9" }

-- Other Income
grid.set{ sheet=2, cell="A18", value="Other Income (Expense)" }
grid.set{ sheet=2, cell="B18", value=333 }
grid.set{ sheet=2, cell="C18", value=788 }
grid.set{ sheet=2, cell="D18", value=1122 }
grid.set{ sheet=2, cell="E18", value=1000 }
grid.set{ sheet=2, cell="F18", value=1000 }
grid.set{ sheet=2, cell="G18", value=1000 }
grid.set{ sheet=2, cell="H18", value=1000 }
grid.set{ sheet=2, cell="I18", value=1000 }

-- Pre-tax Income
grid.set{ sheet=2, cell="A19", value="Income Before Taxes" }
grid.set{ sheet=2, cell="B19", value="=B15+B17+B18" }
grid.set{ sheet=2, cell="C19", value="=C15+C17+C18" }
grid.set{ sheet=2, cell="D19", value="=D15+D17+D18" }
grid.set{ sheet=2, cell="E19", value="=E15+E17+E18" }
grid.set{ sheet=2, cell="F19", value="=F15+F17+F18" }
grid.set{ sheet=2, cell="G19", value="=G15+G17+G18" }
grid.set{ sheet=2, cell="H19", value="=H15+H17+H18" }
grid.set{ sheet=2, cell="I19", value="=I15+I17+I18" }
grid.format{ sheet=2, range="A19:I19", bold=true }

-- Income Tax
grid.set{ sheet=2, cell="A20", value="Income Tax Provision" }
grid.set{ sheet=2, cell="B20", value=-10978 }
grid.set{ sheet=2, cell="C20", value=-16950 }
grid.set{ sheet=2, cell="D20", value=-19651 }
grid.set{ sheet=2, cell="E20", value="=-E19*Assumptions!B8" }
grid.set{ sheet=2, cell="F20", value="=-F19*Assumptions!C8" }
grid.set{ sheet=2, cell="G20", value="=-G19*Assumptions!D8" }
grid.set{ sheet=2, cell="H20", value="=-H19*Assumptions!E8" }
grid.set{ sheet=2, cell="I20", value="=-I19*Assumptions!F8" }

-- Net Income
grid.set{ sheet=2, cell="A22", value="Net Income" }
grid.set{ sheet=2, cell="B22", value="=B19+B20" }
grid.set{ sheet=2, cell="C22", value="=C19+C20" }
grid.set{ sheet=2, cell="D22", value="=D19+D20" }
grid.set{ sheet=2, cell="E22", value="=E19+E20" }
grid.set{ sheet=2, cell="F22", value="=F19+F20" }
grid.set{ sheet=2, cell="G22", value="=G19+G20" }
grid.set{ sheet=2, cell="H22", value="=H19+H20" }
grid.set{ sheet=2, cell="I22", value="=I19+I20" }
grid.format{ sheet=2, range="A22:I22", bold=true }

--------------------------------------------------------------------------------
-- Sheet 3: Balance Sheet
--------------------------------------------------------------------------------
grid.set{ sheet=3, cell="A1", value="Balance Sheet" }
grid.format{ sheet=3, range="A1", bold=true }

-- Year headers
grid.set_batch{ sheet=3, cells={
    {cell="A3", value="($ in millions)"},
    {cell="B3", value="FY2022"},
    {cell="C3", value="FY2023"},
    {cell="D3", value="FY2024"},
    {cell="E3", value="FY2025"},
    {cell="F3", value="FY2026"},
    {cell="G3", value="FY2027"},
    {cell="H3", value="FY2028"},
    {cell="I3", value="FY2029"}
}}
grid.format{ sheet=3, range="A3:I3", bold=true }

-- ASSETS
grid.set{ sheet=3, cell="A5", value="ASSETS" }
grid.format{ sheet=3, range="A5", bold=true }

-- Cash & Equivalents (plug from CF statement)
grid.set{ sheet=3, cell="A6", value="Cash & Equivalents" }
grid.set{ sheet=3, cell="B6", value=13931 }
grid.set{ sheet=3, cell="C6", value=34704 }
grid.set{ sheet=3, cell="D6", value=18315 }
grid.set{ sheet=3, cell="E6", value="='Cash Flow'!E28" }  -- Ending cash from CF
grid.set{ sheet=3, cell="F6", value="='Cash Flow'!F28" }
grid.set{ sheet=3, cell="G6", value="='Cash Flow'!G28" }
grid.set{ sheet=3, cell="H6", value="='Cash Flow'!H28" }
grid.set{ sheet=3, cell="I6", value="='Cash Flow'!I28" }

-- Short-term Investments
grid.set{ sheet=3, cell="A7", value="Short-term Investments" }
grid.set{ sheet=3, cell="B7", value=90826 }
grid.set{ sheet=3, cell="C7", value=76558 }
grid.set{ sheet=3, cell="D7", value=57228 }
grid.set{ sheet=3, cell="E7", value=57228 }
grid.set{ sheet=3, cell="F7", value=57228 }
grid.set{ sheet=3, cell="G7", value=57228 }
grid.set{ sheet=3, cell="H7", value=57228 }
grid.set{ sheet=3, cell="I7", value=57228 }

-- Accounts Receivable
grid.set{ sheet=3, cell="A8", value="Accounts Receivable" }
grid.set{ sheet=3, cell="B8", value=44261 }
grid.set{ sheet=3, cell="C8", value=48688 }
grid.set{ sheet=3, cell="D8", value=56924 }
grid.set{ sheet=3, cell="E8", value="='Income Statement'!E6*(D8/D6)" }
grid.set{ sheet=3, cell="F8", value="='Income Statement'!F6*(D8/D6)" }
grid.set{ sheet=3, cell="G8", value="='Income Statement'!G6*(D8/D6)" }
grid.set{ sheet=3, cell="H8", value="='Income Statement'!H6*(D8/D6)" }
grid.set{ sheet=3, cell="I8", value="='Income Statement'!I6*(D8/D6)" }

-- Other Current Assets
grid.set{ sheet=3, cell="A9", value="Other Current Assets" }
grid.set{ sheet=3, cell="B9", value=16924 }
grid.set{ sheet=3, cell="C9", value=21807 }
grid.set{ sheet=3, cell="D9", value=29553 }
grid.set{ sheet=3, cell="E9", value="='Income Statement'!E6*(D9/D6)" }
grid.set{ sheet=3, cell="F9", value="='Income Statement'!F6*(D9/D6)" }
grid.set{ sheet=3, cell="G9", value="='Income Statement'!G6*(D9/D6)" }
grid.set{ sheet=3, cell="H9", value="='Income Statement'!H6*(D9/D6)" }
grid.set{ sheet=3, cell="I9", value="='Income Statement'!I6*(D9/D6)" }

-- Total Current Assets
grid.set{ sheet=3, cell="A10", value="Total Current Assets" }
grid.set{ sheet=3, cell="B10", value="=SUM(B6:B9)" }
grid.set{ sheet=3, cell="C10", value="=SUM(C6:C9)" }
grid.set{ sheet=3, cell="D10", value="=SUM(D6:D9)" }
grid.set{ sheet=3, cell="E10", value="=SUM(E6:E9)" }
grid.set{ sheet=3, cell="F10", value="=SUM(F6:F9)" }
grid.set{ sheet=3, cell="G10", value="=SUM(G6:G9)" }
grid.set{ sheet=3, cell="H10", value="=SUM(H6:H9)" }
grid.set{ sheet=3, cell="I10", value="=SUM(I6:I9)" }
grid.format{ sheet=3, range="A10:I10", bold=true }

-- PP&E Net
grid.set{ sheet=3, cell="A12", value="Property, Plant & Equipment, Net" }
grid.set{ sheet=3, cell="B12", value=74398 }
grid.set{ sheet=3, cell="C12", value=95641 }
grid.set{ sheet=3, cell="D12", value=135591 }
-- PP&E = Prior PP&E + CapEx - D&A
grid.set{ sheet=3, cell="E12", value="=D12+'Cash Flow'!E12+'Cash Flow'!E8" }
grid.set{ sheet=3, cell="F12", value="=E12+'Cash Flow'!F12+'Cash Flow'!F8" }
grid.set{ sheet=3, cell="G12", value="=F12+'Cash Flow'!G12+'Cash Flow'!G8" }
grid.set{ sheet=3, cell="H12", value="=G12+'Cash Flow'!H12+'Cash Flow'!H8" }
grid.set{ sheet=3, cell="I12", value="=H12+'Cash Flow'!I12+'Cash Flow'!I8" }

-- Goodwill & Intangibles
grid.set{ sheet=3, cell="A13", value="Goodwill & Intangibles" }
grid.set{ sheet=3, cell="B13", value=67524 }
grid.set{ sheet=3, cell="C13", value=67886 }
grid.set{ sheet=3, cell="D13", value=119220 }
grid.set{ sheet=3, cell="E13", value=119220 }
grid.set{ sheet=3, cell="F13", value=119220 }
grid.set{ sheet=3, cell="G13", value=119220 }
grid.set{ sheet=3, cell="H13", value=119220 }
grid.set{ sheet=3, cell="I13", value=119220 }

-- Other Long-term Assets
grid.set{ sheet=3, cell="A14", value="Other Long-term Assets" }
grid.set{ sheet=3, cell="B14", value=21897 }
grid.set{ sheet=3, cell="C14", value=30601 }
grid.set{ sheet=3, cell="D14", value=38545 }
grid.set{ sheet=3, cell="E14", value=38545 }
grid.set{ sheet=3, cell="F14", value=38545 }
grid.set{ sheet=3, cell="G14", value=38545 }
grid.set{ sheet=3, cell="H14", value=38545 }
grid.set{ sheet=3, cell="I14", value=38545 }

-- Total Assets
grid.set{ sheet=3, cell="A15", value="Total Assets" }
grid.set{ sheet=3, cell="B15", value="=B10+B12+B13+B14" }
grid.set{ sheet=3, cell="C15", value="=C10+C12+C13+C14" }
grid.set{ sheet=3, cell="D15", value="=D10+D12+D13+D14" }
grid.set{ sheet=3, cell="E15", value="=E10+E12+E13+E14" }
grid.set{ sheet=3, cell="F15", value="=F10+F12+F13+F14" }
grid.set{ sheet=3, cell="G15", value="=G10+G12+G13+G14" }
grid.set{ sheet=3, cell="H15", value="=H10+H12+H13+H14" }
grid.set{ sheet=3, cell="I15", value="=I10+I12+I13+I14" }
grid.format{ sheet=3, range="A15:I15", bold=true }

-- LIABILITIES
grid.set{ sheet=3, cell="A17", value="LIABILITIES" }
grid.format{ sheet=3, range="A17", bold=true }

-- Accounts Payable
grid.set{ sheet=3, cell="A18", value="Accounts Payable" }
grid.set{ sheet=3, cell="B18", value=19000 }
grid.set{ sheet=3, cell="C18", value=18095 }
grid.set{ sheet=3, cell="D18", value=21996 }
grid.set{ sheet=3, cell="E18", value="='Income Statement'!E6*(D18/D6)" }
grid.set{ sheet=3, cell="F18", value="='Income Statement'!F6*(D18/D6)" }
grid.set{ sheet=3, cell="G18", value="='Income Statement'!G6*(D18/D6)" }
grid.set{ sheet=3, cell="H18", value="='Income Statement'!H6*(D18/D6)" }
grid.set{ sheet=3, cell="I18", value="='Income Statement'!I6*(D18/D6)" }

-- Deferred Revenue & Other
grid.set{ sheet=3, cell="A19", value="Deferred Revenue & Other CL" }
grid.set{ sheet=3, cell="B19", value=76827 }
grid.set{ sheet=3, cell="C19", value=86298 }
grid.set{ sheet=3, cell="D19", value=103369 }
grid.set{ sheet=3, cell="E19", value="='Income Statement'!E6*(D19/D6)" }
grid.set{ sheet=3, cell="F19", value="='Income Statement'!F6*(D19/D6)" }
grid.set{ sheet=3, cell="G19", value="='Income Statement'!G6*(D19/D6)" }
grid.set{ sheet=3, cell="H19", value="='Income Statement'!H6*(D19/D6)" }
grid.set{ sheet=3, cell="I19", value="='Income Statement'!I6*(D19/D6)" }

-- Long-term Debt
grid.set{ sheet=3, cell="A20", value="Long-term Debt" }
grid.set{ sheet=3, cell="B20", value=47032 }
grid.set{ sheet=3, cell="C20", value=41990 }
grid.set{ sheet=3, cell="D20", value=42688 }
grid.set{ sheet=3, cell="E20", value=42688 }
grid.set{ sheet=3, cell="F20", value=42688 }
grid.set{ sheet=3, cell="G20", value=42688 }
grid.set{ sheet=3, cell="H20", value=42688 }
grid.set{ sheet=3, cell="I20", value=42688 }

-- Other Long-term Liabilities
grid.set{ sheet=3, cell="A21", value="Other Long-term Liabilities" }
grid.set{ sheet=3, cell="B21", value=35727 }
grid.set{ sheet=3, cell="C21", value=48676 }
grid.set{ sheet=3, cell="D21", value=53567 }
grid.set{ sheet=3, cell="E21", value=53567 }
grid.set{ sheet=3, cell="F21", value=53567 }
grid.set{ sheet=3, cell="G21", value=53567 }
grid.set{ sheet=3, cell="H21", value=53567 }
grid.set{ sheet=3, cell="I21", value=53567 }

-- Total Liabilities
grid.set{ sheet=3, cell="A22", value="Total Liabilities" }
grid.set{ sheet=3, cell="B22", value="=B18+B19+B20+B21" }
grid.set{ sheet=3, cell="C22", value="=C18+C19+C20+C21" }
grid.set{ sheet=3, cell="D22", value="=D18+D19+D20+D21" }
grid.set{ sheet=3, cell="E22", value="=E18+E19+E20+E21" }
grid.set{ sheet=3, cell="F22", value="=F18+F19+F20+F21" }
grid.set{ sheet=3, cell="G22", value="=G18+G19+G20+G21" }
grid.set{ sheet=3, cell="H22", value="=H18+H19+H20+H21" }
grid.set{ sheet=3, cell="I22", value="=I18+I19+I20+I21" }
grid.format{ sheet=3, range="A22:I22", bold=true }

-- EQUITY
grid.set{ sheet=3, cell="A24", value="SHAREHOLDERS' EQUITY" }
grid.format{ sheet=3, range="A24", bold=true }

-- Retained Earnings (prior + NI - dividends - buybacks)
grid.set{ sheet=3, cell="A25", value="Retained Earnings" }
grid.set{ sheet=3, cell="B25", value=84281 }
grid.set{ sheet=3, cell="C25", value=118848 }
grid.set{ sheet=3, cell="D25", value=115440 }
grid.set{ sheet=3, cell="E25", value="=D25+'Income Statement'!E22-Assumptions!B13-Assumptions!B14" }
grid.set{ sheet=3, cell="F25", value="=E25+'Income Statement'!F22-Assumptions!C13-Assumptions!C14" }
grid.set{ sheet=3, cell="G25", value="=F25+'Income Statement'!G22-Assumptions!D13-Assumptions!D14" }
grid.set{ sheet=3, cell="H25", value="=G25+'Income Statement'!H22-Assumptions!E13-Assumptions!E14" }
grid.set{ sheet=3, cell="I25", value="=H25+'Income Statement'!I22-Assumptions!F13-Assumptions!F14" }

-- Other Equity (AOCI, etc.) - kept flat
grid.set{ sheet=3, cell="A26", value="Other Equity Components" }
grid.set{ sheet=3, cell="B26", value=8374 }
grid.set{ sheet=3, cell="C26", value=-6611 }
grid.set{ sheet=3, cell="D26", value=1728 }
grid.set{ sheet=3, cell="E26", value=1728 }
grid.set{ sheet=3, cell="F26", value=1728 }
grid.set{ sheet=3, cell="G26", value=1728 }
grid.set{ sheet=3, cell="H26", value=1728 }
grid.set{ sheet=3, cell="I26", value=1728 }

-- Total Equity
grid.set{ sheet=3, cell="A27", value="Total Shareholders' Equity" }
grid.set{ sheet=3, cell="B27", value="=B25+B26" }
grid.set{ sheet=3, cell="C27", value="=C25+C26" }
grid.set{ sheet=3, cell="D27", value="=D25+D26" }
grid.set{ sheet=3, cell="E27", value="=E25+E26" }
grid.set{ sheet=3, cell="F27", value="=F25+F26" }
grid.set{ sheet=3, cell="G27", value="=G25+G26" }
grid.set{ sheet=3, cell="H27", value="=H25+H26" }
grid.set{ sheet=3, cell="I27", value="=I25+I26" }
grid.format{ sheet=3, range="A27:I27", bold=true }

-- Total Liabilities + Equity
grid.set{ sheet=3, cell="A29", value="Total Liabilities + Equity" }
grid.set{ sheet=3, cell="B29", value="=B22+B27" }
grid.set{ sheet=3, cell="C29", value="=C22+C27" }
grid.set{ sheet=3, cell="D29", value="=D22+D27" }
grid.set{ sheet=3, cell="E29", value="=E22+E27" }
grid.set{ sheet=3, cell="F29", value="=F22+F27" }
grid.set{ sheet=3, cell="G29", value="=G22+G27" }
grid.set{ sheet=3, cell="H29", value="=H22+H27" }
grid.set{ sheet=3, cell="I29", value="=I22+I27" }
grid.format{ sheet=3, range="A29:I29", bold=true }

-- Balance Check
grid.set{ sheet=3, cell="A31", value="CHECK: Assets - (L+E)" }
grid.set{ sheet=3, cell="B31", value="=B15-B29" }
grid.set{ sheet=3, cell="C31", value="=C15-C29" }
grid.set{ sheet=3, cell="D31", value="=D15-D29" }
grid.set{ sheet=3, cell="E31", value="=E15-E29" }
grid.set{ sheet=3, cell="F31", value="=F15-F29" }
grid.set{ sheet=3, cell="G31", value="=G15-G29" }
grid.set{ sheet=3, cell="H31", value="=H15-H29" }
grid.set{ sheet=3, cell="I31", value="=I15-I29" }
grid.format{ sheet=3, range="A31:I31", bold=true }

grid.set{ sheet=3, cell="A32", value="Balance Check (must = 0)" }
grid.set{ sheet=3, cell="B32", value="=IF(ABS(B31)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=3, cell="C32", value="=IF(ABS(C31)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=3, cell="D32", value="=IF(ABS(D31)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=3, cell="E32", value="=IF(ABS(E31)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=3, cell="F32", value="=IF(ABS(F31)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=3, cell="G32", value="=IF(ABS(G31)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=3, cell="H32", value="=IF(ABS(H31)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=3, cell="I32", value="=IF(ABS(I31)<1,\"OK\",\"ERROR\")" }

--------------------------------------------------------------------------------
-- Sheet 4: Cash Flow Statement
--------------------------------------------------------------------------------
grid.set{ sheet=4, cell="A1", value="Cash Flow Statement" }
grid.format{ sheet=4, range="A1", bold=true }

-- Year headers
grid.set_batch{ sheet=4, cells={
    {cell="A3", value="($ in millions)"},
    {cell="B3", value="FY2022"},
    {cell="C3", value="FY2023"},
    {cell="D3", value="FY2024"},
    {cell="E3", value="FY2025"},
    {cell="F3", value="FY2026"},
    {cell="G3", value="FY2027"},
    {cell="H3", value="FY2028"},
    {cell="I3", value="FY2029"}
}}
grid.format{ sheet=4, range="A3:I3", bold=true }

-- OPERATING ACTIVITIES
grid.set{ sheet=4, cell="A5", value="OPERATING ACTIVITIES" }
grid.format{ sheet=4, range="A5", bold=true }

-- Net Income (from IS)
grid.set{ sheet=4, cell="A6", value="Net Income" }
grid.set{ sheet=4, cell="B6", value="='Income Statement'!B22" }
grid.set{ sheet=4, cell="C6", value="='Income Statement'!C22" }
grid.set{ sheet=4, cell="D6", value="='Income Statement'!D22" }
grid.set{ sheet=4, cell="E6", value="='Income Statement'!E22" }
grid.set{ sheet=4, cell="F6", value="='Income Statement'!F22" }
grid.set{ sheet=4, cell="G6", value="='Income Statement'!G22" }
grid.set{ sheet=4, cell="H6", value="='Income Statement'!H22" }
grid.set{ sheet=4, cell="I6", value="='Income Statement'!I22" }

-- D&A (add back)
grid.set{ sheet=4, cell="A8", value="Depreciation & Amortization" }
grid.set{ sheet=4, cell="B8", value=14460 }
grid.set{ sheet=4, cell="C8", value=13861 }
grid.set{ sheet=4, cell="D8", value=22287 }
grid.set{ sheet=4, cell="E8", value="='Income Statement'!E6*Assumptions!B10" }
grid.set{ sheet=4, cell="F8", value="='Income Statement'!F6*Assumptions!C10" }
grid.set{ sheet=4, cell="G8", value="='Income Statement'!G6*Assumptions!D10" }
grid.set{ sheet=4, cell="H8", value="='Income Statement'!H6*Assumptions!E10" }
grid.set{ sheet=4, cell="I8", value="='Income Statement'!I6*Assumptions!F10" }

-- Change in Working Capital
grid.set{ sheet=4, cell="A9", value="Change in Working Capital" }
grid.set{ sheet=4, cell="B9", value=-929 }
grid.set{ sheet=4, cell="C9", value=3280 }
grid.set{ sheet=4, cell="D9", value=-5523 }
-- NWC change = (Current NWC % - Prior NWC %) * Revenue
grid.set{ sheet=4, cell="E9", value="=-('Income Statement'!E6*Assumptions!B12-'Income Statement'!D6*(-0.15))" }
grid.set{ sheet=4, cell="F9", value="=-('Income Statement'!F6*Assumptions!C12-'Income Statement'!E6*Assumptions!B12)" }
grid.set{ sheet=4, cell="G9", value="=-('Income Statement'!G6*Assumptions!D12-'Income Statement'!F6*Assumptions!C12)" }
grid.set{ sheet=4, cell="H9", value="=-('Income Statement'!H6*Assumptions!E12-'Income Statement'!G6*Assumptions!D12)" }
grid.set{ sheet=4, cell="I9", value="=-('Income Statement'!I6*Assumptions!F12-'Income Statement'!H6*Assumptions!E12)" }

-- Other Operating
grid.set{ sheet=4, cell="A10", value="Other Operating Adjustments" }
grid.set{ sheet=4, cell="B10", value=4139 }
grid.set{ sheet=4, cell="C10", value=11832 }
grid.set{ sheet=4, cell="D10", value=-3536 }
grid.set{ sheet=4, cell="E10", value=0 }
grid.set{ sheet=4, cell="F10", value=0 }
grid.set{ sheet=4, cell="G10", value=0 }
grid.set{ sheet=4, cell="H10", value=0 }
grid.set{ sheet=4, cell="I10", value=0 }

-- Cash from Operations
grid.set{ sheet=4, cell="A11", value="Cash from Operations" }
grid.set{ sheet=4, cell="B11", value="=B6+B8+B9+B10" }
grid.set{ sheet=4, cell="C11", value="=C6+C8+C9+C10" }
grid.set{ sheet=4, cell="D11", value="=D6+D8+D9+D10" }
grid.set{ sheet=4, cell="E11", value="=E6+E8+E9+E10" }
grid.set{ sheet=4, cell="F11", value="=F6+F8+F9+F10" }
grid.set{ sheet=4, cell="G11", value="=G6+G8+G9+G10" }
grid.set{ sheet=4, cell="H11", value="=H6+H8+H9+H10" }
grid.set{ sheet=4, cell="I11", value="=I6+I8+I9+I10" }
grid.format{ sheet=4, range="A11:I11", bold=true }

-- INVESTING ACTIVITIES
grid.set{ sheet=4, cell="A13", value="INVESTING ACTIVITIES" }
grid.format{ sheet=4, range="A13", bold=true }

-- CapEx
grid.set{ sheet=4, cell="A12", value="Capital Expenditures" }
grid.set{ sheet=4, cell="B12", value=-23886 }
grid.set{ sheet=4, cell="C12", value=-28107 }
grid.set{ sheet=4, cell="D12", value=-44477 }
grid.set{ sheet=4, cell="E12", value="=-'Income Statement'!E6*Assumptions!B11" }
grid.set{ sheet=4, cell="F12", value="=-'Income Statement'!F6*Assumptions!C11" }
grid.set{ sheet=4, cell="G12", value="=-'Income Statement'!G6*Assumptions!D11" }
grid.set{ sheet=4, cell="H12", value="=-'Income Statement'!H6*Assumptions!E11" }
grid.set{ sheet=4, cell="I12", value="=-'Income Statement'!I6*Assumptions!F11" }

-- Other Investing
grid.set{ sheet=4, cell="A14", value="Acquisitions & Other" }
grid.set{ sheet=4, cell="B14", value=-33147 }
grid.set{ sheet=4, cell="C14", value=6226 }
grid.set{ sheet=4, cell="D14", value=1824 }
grid.set{ sheet=4, cell="E14", value=0 }
grid.set{ sheet=4, cell="F14", value=0 }
grid.set{ sheet=4, cell="G14", value=0 }
grid.set{ sheet=4, cell="H14", value=0 }
grid.set{ sheet=4, cell="I14", value=0 }

-- Cash from Investing
grid.set{ sheet=4, cell="A15", value="Cash from Investing" }
grid.set{ sheet=4, cell="B15", value="=B12+B14" }
grid.set{ sheet=4, cell="C15", value="=C12+C14" }
grid.set{ sheet=4, cell="D15", value="=D12+D14" }
grid.set{ sheet=4, cell="E15", value="=E12+E14" }
grid.set{ sheet=4, cell="F15", value="=F12+F14" }
grid.set{ sheet=4, cell="G15", value="=G12+G14" }
grid.set{ sheet=4, cell="H15", value="=H12+H14" }
grid.set{ sheet=4, cell="I15", value="=I12+I14" }
grid.format{ sheet=4, range="A15:I15", bold=true }

-- FINANCING ACTIVITIES
grid.set{ sheet=4, cell="A17", value="FINANCING ACTIVITIES" }
grid.format{ sheet=4, range="A17", bold=true }

-- Dividends
grid.set{ sheet=4, cell="A18", value="Dividends Paid" }
grid.set{ sheet=4, cell="B18", value=-18135 }
grid.set{ sheet=4, cell="C18", value=-19800 }
grid.set{ sheet=4, cell="D18", value=-21771 }
grid.set{ sheet=4, cell="E18", value="=-Assumptions!B13" }
grid.set{ sheet=4, cell="F18", value="=-Assumptions!C13" }
grid.set{ sheet=4, cell="G18", value="=-Assumptions!D13" }
grid.set{ sheet=4, cell="H18", value="=-Assumptions!E13" }
grid.set{ sheet=4, cell="I18", value="=-Assumptions!F13" }

-- Share Repurchases
grid.set{ sheet=4, cell="A19", value="Share Repurchases" }
grid.set{ sheet=4, cell="B19", value=-32696 }
grid.set{ sheet=4, cell="C19", value=-22245 }
grid.set{ sheet=4, cell="D19", value=-17254 }
grid.set{ sheet=4, cell="E19", value="=-Assumptions!B14" }
grid.set{ sheet=4, cell="F19", value="=-Assumptions!C14" }
grid.set{ sheet=4, cell="G19", value="=-Assumptions!D14" }
grid.set{ sheet=4, cell="H19", value="=-Assumptions!E14" }
grid.set{ sheet=4, cell="I19", value="=-Assumptions!F14" }

-- Debt Changes
grid.set{ sheet=4, cell="A20", value="Debt Issued / (Repaid)" }
grid.set{ sheet=4, cell="B20", value=-9023 }
grid.set{ sheet=4, cell="C20", value=-2750 }
grid.set{ sheet=4, cell="D20", value=698 }
grid.set{ sheet=4, cell="E20", value=0 }
grid.set{ sheet=4, cell="F20", value=0 }
grid.set{ sheet=4, cell="G20", value=0 }
grid.set{ sheet=4, cell="H20", value=0 }
grid.set{ sheet=4, cell="I20", value=0 }

-- Other Financing
grid.set{ sheet=4, cell="A21", value="Other Financing" }
grid.set{ sheet=4, cell="B21", value=1822 }
grid.set{ sheet=4, cell="C21", value=1866 }
grid.set{ sheet=4, cell="D21", value=-5765 }
grid.set{ sheet=4, cell="E21", value=0 }
grid.set{ sheet=4, cell="F21", value=0 }
grid.set{ sheet=4, cell="G21", value=0 }
grid.set{ sheet=4, cell="H21", value=0 }
grid.set{ sheet=4, cell="I21", value=0 }

-- Cash from Financing
grid.set{ sheet=4, cell="A22", value="Cash from Financing" }
grid.set{ sheet=4, cell="B22", value="=B18+B19+B20+B21" }
grid.set{ sheet=4, cell="C22", value="=C18+C19+C20+C21" }
grid.set{ sheet=4, cell="D22", value="=D18+D19+D20+D21" }
grid.set{ sheet=4, cell="E22", value="=E18+E19+E20+E21" }
grid.set{ sheet=4, cell="F22", value="=F18+F19+F20+F21" }
grid.set{ sheet=4, cell="G22", value="=G18+G19+G20+G21" }
grid.set{ sheet=4, cell="H22", value="=H18+H19+H20+H21" }
grid.set{ sheet=4, cell="I22", value="=I18+I19+I20+I21" }
grid.format{ sheet=4, range="A22:I22", bold=true }

-- FX Effect
grid.set{ sheet=4, cell="A24", value="Effect of FX on Cash" }
grid.set{ sheet=4, cell="B24", value=-141 }
grid.set{ sheet=4, cell="C24", value=-194 }
grid.set{ sheet=4, cell="D24", value=-170 }
grid.set{ sheet=4, cell="E24", value=0 }
grid.set{ sheet=4, cell="F24", value=0 }
grid.set{ sheet=4, cell="G24", value=0 }
grid.set{ sheet=4, cell="H24", value=0 }
grid.set{ sheet=4, cell="I24", value=0 }

-- Net Change in Cash
grid.set{ sheet=4, cell="A26", value="Net Change in Cash" }
grid.set{ sheet=4, cell="B26", value="=B11+B15+B22+B24" }
grid.set{ sheet=4, cell="C26", value="=C11+C15+C22+C24" }
grid.set{ sheet=4, cell="D26", value="=D11+D15+D22+D24" }
grid.set{ sheet=4, cell="E26", value="=E11+E15+E22+E24" }
grid.set{ sheet=4, cell="F26", value="=F11+F15+F22+F24" }
grid.set{ sheet=4, cell="G26", value="=G11+G15+G22+G24" }
grid.set{ sheet=4, cell="H26", value="=H11+H15+H22+H24" }
grid.set{ sheet=4, cell="I26", value="=I11+I15+I22+I24" }
grid.format{ sheet=4, range="A26:I26", bold=true }

-- Beginning Cash
grid.set{ sheet=4, cell="A27", value="Beginning Cash" }
grid.set{ sheet=4, cell="B27", value=14224 }
grid.set{ sheet=4, cell="C27", value="=B28" }
grid.set{ sheet=4, cell="D27", value="=C28" }
grid.set{ sheet=4, cell="E27", value="=D28" }
grid.set{ sheet=4, cell="F27", value="=E28" }
grid.set{ sheet=4, cell="G27", value="=F28" }
grid.set{ sheet=4, cell="H27", value="=G28" }
grid.set{ sheet=4, cell="I27", value="=H28" }

-- Ending Cash
grid.set{ sheet=4, cell="A28", value="Ending Cash" }
grid.set{ sheet=4, cell="B28", value="=B27+B26" }
grid.set{ sheet=4, cell="C28", value="=C27+C26" }
grid.set{ sheet=4, cell="D28", value="=D27+D26" }
grid.set{ sheet=4, cell="E28", value="=E27+E26" }
grid.set{ sheet=4, cell="F28", value="=F27+F26" }
grid.set{ sheet=4, cell="G28", value="=G27+G26" }
grid.set{ sheet=4, cell="H28", value="=H27+H26" }
grid.set{ sheet=4, cell="I28", value="=I27+I26" }
grid.format{ sheet=4, range="A28:I28", bold=true }

-- Cash Check
grid.set{ sheet=4, cell="A30", value="CHECK: CF Ending Cash - BS Cash" }
grid.set{ sheet=4, cell="B30", value="=B28-'Balance Sheet'!B6" }
grid.set{ sheet=4, cell="C30", value="=C28-'Balance Sheet'!C6" }
grid.set{ sheet=4, cell="D30", value="=D28-'Balance Sheet'!D6" }
grid.set{ sheet=4, cell="E30", value="=E28-'Balance Sheet'!E6" }
grid.set{ sheet=4, cell="F30", value="=F28-'Balance Sheet'!F6" }
grid.set{ sheet=4, cell="G30", value="=G28-'Balance Sheet'!G6" }
grid.set{ sheet=4, cell="H30", value="=H28-'Balance Sheet'!H6" }
grid.set{ sheet=4, cell="I30", value="=I28-'Balance Sheet'!I6" }
grid.format{ sheet=4, range="A30:I30", bold=true }

grid.set{ sheet=4, cell="A31", value="Cash Check (must = 0)" }
grid.set{ sheet=4, cell="B31", value="=IF(ABS(B30)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=4, cell="C31", value="=IF(ABS(C30)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=4, cell="D31", value="=IF(ABS(D30)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=4, cell="E31", value="=IF(ABS(E30)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=4, cell="F31", value="=IF(ABS(F30)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=4, cell="G31", value="=IF(ABS(G30)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=4, cell="H31", value="=IF(ABS(H30)<1,\"OK\",\"ERROR\")" }
grid.set{ sheet=4, cell="I31", value="=IF(ABS(I30)<1,\"OK\",\"ERROR\")" }
