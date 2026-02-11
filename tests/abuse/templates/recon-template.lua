-- Recon template builder
-- Builds a 2-sheet workbook: tx (data target) + summary (formulas)
--
-- tx sheet: headers in row 1, data filled by `vgrid fill --target tx!A2`
-- summary sheet: SUMIF aggregations + variance check cell at B7

-- ── tx sheet (Sheet1 = default) ──
-- Headers only; data rows filled by `vgrid fill`
set("A1", "effective_date")
set("B1", "posted_date")
set("C1", "amount_minor")
set("D1", "currency")
set("E1", "type")
set("F1", "source")
set("G1", "source_id")
set("H1", "group_id")
set("I1", "description")
set("J1", "amount")

-- ── summary sheet (Sheet2) ──
-- Switch to summary sheet
sheet("summary")

-- Labels
set("A1", "Category")
set("B1", "Total (minor units)")

set("A2", "Charges")
set("A3", "Payouts")
set("A4", "Fees")
set("A5", "Refunds")
set("A6", "Adjustments")
set("A7", "Variance")

-- SUMIF formulas referencing tx sheet
-- Sum amount_minor (col C) where type (col E) matches
set("B2", "=SUMIF(Sheet1!E:E,\"charge\",Sheet1!C:C)")
set("B3", "=SUMIF(Sheet1!E:E,\"payout\",Sheet1!C:C)")
set("B4", "=SUMIF(Sheet1!E:E,\"fee\",Sheet1!C:C)")
set("B5", "=SUMIF(Sheet1!E:E,\"refund\",Sheet1!C:C)")
set("B6", "=SUMIF(Sheet1!E:E,\"adjustment\",Sheet1!C:C)")

-- Variance: charges + fees + refunds + adjustments + payouts should net to 0
-- (payouts = negative of charges minus fees minus refunds)
-- This is the cell we assert: summary!B7 should be 0 (or within tolerance)
set("B7", "=B2+B3+B4+B5+B6")
