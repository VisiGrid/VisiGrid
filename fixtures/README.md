# VisiGrid Demo Fixtures

Example files demonstrating VisiGrid's explainability features.

## explainability-demo.csv

A simple sales forecast model with multi-level dependencies.

**To use:**

1. Open VisiGrid
2. File > Open > select `explainability-demo.csv`
3. Press F9 to enable Verified Mode
4. Click on cell E15 (After Tax profit)
5. Press F1 to open Inspector

**What you'll see:**

- **Trust Header**: Shows depth and impact (affects 0 cells, max depth 4)
- **Inputs/Outputs**: Click cells in the DAG to trace the path
- **Proof Section**: See evaluation order and recompute timestamps

**Try these actions:**

1. **Edit an input** (B5 = Revenue Per Unit): Watch the status bar flash "Verified" as all dependents recompute
2. **Path trace**: Click E5 (Total Revenue), then click any input to see the data flow
3. **Create a cycle**: Try typing `=E15` in cell B5. VisiGrid will reject it with a cycle error.
4. **History panel**: Press Ctrl+Shift+Y to see provenance for your edits

## Creating Named Ranges

After import, define some named ranges to test Named Range Intelligence:

1. Select B5:B7 (the input cells)
2. Press Ctrl+Shift+N to open the Names panel
3. Click "+" to create a new named range called "Inputs"
4. Select the range and press F1 to see usage count and depth

## Verified Mode Semantics

With Verified Mode on (F9):
- Status bar shows "Verified" when all values are current
- Any edit triggers a full topological recompute
- Cycles are detected at edit-time (not hidden as #VALUE)
