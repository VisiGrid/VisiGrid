# VisiGrid Agent Prompts

Three quality prompts for agent-driven spreadsheet builds. Each follows the workflow:
**write Lua → apply → inspect → verify**

---

## Prompt 1: Build a 12-Month Cashflow Model

```
Build a 12-month cashflow model in cashflow.sheet with:

Inputs (tagged with meta for identification):
- Starting cash: $50,000
- Monthly revenue: $15,000
- Monthly expenses: $12,000
- Growth rate: 3% per month

Model structure:
- Row 1: Headers (Month, Revenue, Expenses, Net, Cash Balance)
- Rows 2-13: Monthly calculations
- Row 14: Annual totals

Requirements:
1. Revenue grows by the growth rate each month
2. Expenses are fixed
3. Net = Revenue - Expenses
4. Cash Balance = Previous Cash + Net
5. Format headers and totals as bold (style only)
6. Tag input cells with meta({ type = "input" })

Workflow:
1. Write the Lua build script
2. Run: vgrid sheet apply cashflow.sheet --lua build.lua --json
3. Inspect B2 (first revenue) and N2 (first cash balance) to verify formulas
4. Inspect B14 (total revenue) to verify the sum
5. Run: vgrid sheet fingerprint cashflow.sheet --json
6. Report the fingerprint for future verification
```

---

## Prompt 2: Normalize CSV to Standard Schema

```
I have this raw CSV export:

```csv
date,vendor_name,amt,category
2024-01-15,AMAZON WEB SERVICES,1234.56,Cloud
2024-01-16,GOOGLE WORKSPACE,-99.00,SaaS
2024-01-17,STRIPE INC,(500.00),Payments
```

Normalize it into a .sheet with this standard schema:
- Date (ISO format: YYYY-MM-DD)
- Vendor (cleaned: title case, no "INC" suffix)
- Amount (positive number, absolute value)
- Direction (Debit/Credit based on original sign)
- Category (unchanged)

Requirements:
1. Parse the CSV and build the normalized sheet
2. Handle parentheses as negative (accounting format)
3. Tag the header row with meta({ role = "header" })
4. Add a Totals row at the bottom with =SUM for Amount

Workflow:
1. Write a Lua script that sets each normalized cell
2. Run: vgrid sheet apply normalized.sheet --lua normalize.lua --json
3. Inspect A2:E4 to verify the normalized data
4. Inspect E5 to verify the total formula
5. Report the fingerprint
```

---

## Prompt 3: Modify Existing Sheet and Verify Change

```
Given the existing model.sheet with fingerprint v1:48:abc123..., add a totals row and verify that only the expected cells changed.

Task:
1. First, inspect the current state:
   vgrid sheet inspect model.sheet --json

2. Identify the last data row (let's say it's row 19)

3. Add a totals row at row 21:
   - A21: "TOTAL"
   - B21: =SUM(B8:B19)
   - C21: =C19 (final cumulative)
   - Tag with meta({ role = "total" })
   - Style with bold

4. Build by writing a Lua script that:
   - Copies all existing data (or reads and re-sets)
   - Adds the new totals row

5. Apply and get new fingerprint:
   vgrid sheet apply model_v2.sheet --lua build_v2.lua --json

6. Verify the change:
   - Old fingerprint: v1:48:abc123...
   - New fingerprint should differ (we added cells)
   - Inspect B21 to confirm the total formula is correct

Report:
- Before fingerprint
- After fingerprint
- The specific cells that were added
- Confirmation that the total formula evaluates correctly
```

---

## Prompt 4: Build a Multi-Sheet Workbook

```
Build a Q1 financial workbook with three sheets:

Sheet 1 (Summary):
- A1: Title "Q1 Financial Summary"
- A4:B4: Total Revenue (formula referencing Sheet2)
- A5:B5: Total Expenses (formula referencing Sheet3)
- A6:B6: Net Income (Revenue - Expenses)

Sheet 2 (Revenue):
- Revenue by category (Product Sales, Services, Subscriptions, Licensing)
- Each with an amount
- Total row with =SUM formula

Sheet 3 (Expenses):
- Expenses by category (Salaries, Infrastructure, Marketing, Operations)
- Each with an amount
- Total row with =SUM formula

Requirements:
1. Use grid.set{ sheet=N, ... } for multi-sheet writes (N is 1-indexed)
2. Use grid.set_batch for efficient bulk updates
3. Cross-sheet formulas use Sheet2!B:B syntax
4. Headers and totals should be bold

Workflow:
1. Write the Lua build script using grid.* API
2. Run: vgrid sheet apply workbook.sheet --lua build.lua --json
3. Inspect Sheet1 B4 to verify cross-sheet formula
4. Inspect Sheet2 B9 and Sheet3 B9 to verify totals
5. Report the fingerprint
```

---

## Multi-Sheet Lua API Reference

The default API (`set`, `meta`, `style`) operates on Sheet1. For multi-sheet workbooks, use the `grid.*` namespace:

```lua
-- Single cell write (sheet is 1-indexed)
grid.set{ sheet=2, cell="A1", value="Hello" }

-- Batch write (more efficient for multiple cells)
grid.set_batch{ sheet=2, cells={
    {cell="A1", value="Region"},
    {cell="B1", value="Amount"},
    {cell="A2", value="North"},
    {cell="B2", value=5000}
}}

-- Formatting (style only, excluded from fingerprint)
grid.format{ sheet=2, range="A1:B1", bold=true }
grid.format{ sheet=3, range="A1", italic=true, underline=true }

-- Cross-sheet formulas
grid.set{ sheet=1, cell="B4", value="=SUM(Sheet2!B:B)" }
grid.set{ sheet=1, cell="B5", value="=Sheet3!B9" }
```

**Key points:**
- Sheets are 1-indexed (sheet=1, sheet=2, etc.)
- Sheets are auto-created as needed — referencing sheet=3 creates Sheet2 and Sheet3 if missing
- Default API (`set`, `meta`, `style`) always targets Sheet1
- Cross-sheet references use `SheetName!CellRef` syntax

---

## Semantic Cell Styles (Interactive Console)

The Lua console (`sheet:*` API) supports semantic cell styles. These convey **meaning**, not formatting — the theme resolves styles to colors. Agents should use these instead of painting explicit fill/font/borders.

```lua
-- Preferred: semantic styling
sheet:input("B2:D10")
sheet:total("A12:F12")
sheet:error("E7")
sheet:warning("C3:C5")
sheet:success("D2")
sheet:note("A1")

-- General method (string name or constant)
sheet:style("A1:C5", "Error")
sheet:style("A1:C5", styles.Warning)

-- Clear back to default
sheet:clear_style("A1:C5")
```

**Available styles:** `Error`, `Warning`, `Success`, `Input`, `Total`, `Note`, `Default`

**Constants table:** `styles.Error` (1), `styles.Warning` (2), `styles.Success` (3), `styles.Input` (4), `styles.Total` (5), `styles.Note` (6), `styles.Default` (0)

**String aliases:** `"warn"` = Warning, `"ok"` = Success, `"totals"` = Total, `"none"` / `"clear"` = Default

**Key rule:** `sheet:style()` sets `cell_style` only. It does NOT paint explicit fill/font/borders. Style is the base layer — the theme resolves it to visual properties.

---

## Key Rules for All Prompts

1. **Always use `--json` output** for parsing tool results
2. **Never assume values** — always inspect after apply
3. **Always capture fingerprint** for audit trail
4. **Style doesn't affect fingerprint** — format freely
5. **Meta does affect fingerprint** — use it for semantic tagging
6. **Use grid.* for multi-sheet** — default API only writes to Sheet1
7. **Prefer semantic styles over formatting** — `sheet:error("A1")` over `style("A1", { bg = "red" })`
