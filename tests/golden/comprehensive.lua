-- api=v1
-- VisiGrid Provenance Script
-- Comprehensive test of all supported operations
-- Expected fingerprint: v1:10:ab2618d9432569ac55470c7e5c09d95f
-- Generated: 2025-01-26T12:00:00Z
-- Actions: 10

-- #1 Set headers in batch
grid.set_batch{ sheet=1, cells={
    { cell="A1", value="Product" },
    { cell="B1", value="Price" },
    { cell="C1", value="Quantity" },
    { cell="D1", value="Total" },
  } }

-- #2 Set data row 1
grid.set_batch{ sheet=1, cells={
    { cell="A2", value="Widget" },
    { cell="B2", value="10.99" },
    { cell="C2", value="5" },
    { cell="D2", value="=B2*C2" },
  } }

-- #3 Set data row 2
grid.set_batch{ sheet=1, cells={
    { cell="A3", value="Gadget" },
    { cell="B3", value="24.99" },
    { cell="C3", value="3" },
    { cell="D3", value="=B3*C3" },
  } }

-- #4 Format headers as bold
grid.format{ sheet=1, range="A1:D1", bold=true }

-- #5 Insert a new row
grid.insert_rows{ sheet=1, at=4, count=1 }

-- #6 Add total row
grid.set{ sheet=1, cell="A4", value="TOTAL" }

-- #7 Add sum formula
grid.set{ sheet=1, cell="D4", value="=SUM(D2:D3)" }

-- #8 Define a named range
grid.define_name{ name="Totals", sheet=1, range="D2:D4" }

-- #9 Set description
grid.set_name_description{ name="Totals", description="Product totals column" }

-- #10 Sort by product name (hashed for fingerprint, not applied to data)
grid.sort{ sheet=1, col=1, ascending=true }
