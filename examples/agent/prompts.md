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
2. Run: visigrid-cli sheet apply cashflow.sheet --lua build.lua --json
3. Inspect B2 (first revenue) and N2 (first cash balance) to verify formulas
4. Inspect B14 (total revenue) to verify the sum
5. Run: visigrid-cli sheet fingerprint cashflow.sheet --json
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
2. Run: visigrid-cli sheet apply normalized.sheet --lua normalize.lua --json
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
   visigrid-cli sheet inspect model.sheet --json

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
   visigrid-cli sheet apply model_v2.sheet --lua build_v2.lua --json

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

## Key Rules for All Prompts

1. **Always use `--json` output** for parsing tool results
2. **Never assume values** — always inspect after apply
3. **Always capture fingerprint** for audit trail
4. **Style doesn't affect fingerprint** — format freely
5. **Meta does affect fingerprint** — use it for semantic tagging
