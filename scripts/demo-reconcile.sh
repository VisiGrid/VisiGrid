#!/usr/bin/env bash
#
# demo-reconcile.sh — Build a reconcile-demo.sheet with scripts and run records.
#
# Demonstrates VisiGrid's audit-grade scripting:
#   1. Import a small dataset
#   2. Run two scripts (sum + flag) with provenance tracking
#   3. Verify run record integrity
#
# Usage: ./scripts/demo-reconcile.sh [output_dir]
#
set -euo pipefail

DIR="${1:-/tmp/vgrid-demo}"
VGRID="${VGRID:-vgrid}"

mkdir -p "$DIR"
SHEET="$DIR/reconcile-demo.sheet"
SCRIPTS_DIR="$DIR/.visigrid/scripts"
mkdir -p "$SCRIPTS_DIR"

echo "=== VisiGrid Reconcile Demo ==="
echo ""

# --- Step 1: Create source data ---
cat > "$DIR/invoices.csv" << 'CSV'
Invoice,Amount,Tax,Total
INV-001,1000,100,
INV-002,2500,250,
INV-003,750,75,
INV-004,3200,320,
INV-005,1800,180,
CSV

echo "1. Created invoices.csv (5 rows, Total column empty)"

# --- Step 2: Import to .sheet ---
$VGRID sheet import "$DIR/invoices.csv" "$SHEET" --headers
echo "2. Imported to $SHEET"
echo ""

# --- Step 3: Create scripts ---
cat > "$SCRIPTS_DIR/sum_total.lua" << 'LUA'
-- Sum Amount + Tax into Total column (D)
for r = 1, sheet:rows() do
    local amount = tonumber(sheet:get_value(r, 2)) or 0
    local tax = tonumber(sheet:get_value(r, 3)) or 0
    if amount > 0 then
        sheet:set_value(r, 4, amount + tax)
    end
end
LUA

cat > "$SCRIPTS_DIR/sum_total.json" << 'JSON'
{
    "schema_version": 1,
    "name": "sum_total",
    "description": "Sum Amount + Tax into Total column",
    "capabilities": ["sheet_read", "sheet_write_values"],
    "author": "demo",
    "version": "1.0.0"
}
JSON

cat > "$SCRIPTS_DIR/flag_large.lua" << 'LUA'
-- Flag invoices over 3000 in column E
for r = 1, sheet:rows() do
    local total = tonumber(sheet:get_value(r, 4)) or 0
    if total > 3000 then
        sheet:set_value(r, 5, "REVIEW")
    end
end
LUA

cat > "$SCRIPTS_DIR/flag_large.json" << 'JSON'
{
    "schema_version": 1,
    "name": "flag_large",
    "description": "Flag invoices with Total > 3000 for review",
    "capabilities": ["sheet_read", "sheet_write_values"],
    "author": "demo",
    "version": "1.0.0"
}
JSON

echo "3. Created 2 project scripts:"
echo "   - sum_total.lua  (Amount + Tax → Total)"
echo "   - flag_large.lua (Total > 3000 → REVIEW)"
echo ""

# --- Step 4: List available scripts ---
echo "4. Available scripts:"
$VGRID scripts list --file "$SHEET"
echo ""

# --- Step 5: Run sum_total (plan first, then apply) ---
echo "5. Dry run sum_total:"
$VGRID scripts run sum_total "$SHEET" --plan
echo ""

echo "6. Apply sum_total:"
$VGRID scripts run sum_total "$SHEET" --apply
echo ""

# --- Step 6: Run flag_large ---
echo "7. Apply flag_large:"
$VGRID scripts run flag_large "$SHEET" --apply
echo ""

# --- Step 7: List run records ---
echo "8. Run records:"
$VGRID runs list "$SHEET"
echo ""

# --- Step 8: Verify integrity ---
echo "9. Verify run record integrity:"
$VGRID runs verify "$SHEET"
echo ""

# --- Step 9: Show final state ---
echo "10. Final sheet state:"
$VGRID sheet peek "$SHEET" 2>/dev/null || echo "    (use 'vgrid sheet peek $SHEET' to inspect)"
echo ""

echo "=== Demo complete ==="
echo ""
echo "Files created:"
echo "  $SHEET"
echo "  $SCRIPTS_DIR/sum_total.lua"
echo "  $SCRIPTS_DIR/flag_large.lua"
echo ""
echo "Try these commands:"
echo "  $VGRID runs list $SHEET --json"
echo "  $VGRID runs verify $SHEET --json"
echo "  $VGRID sheet verify $SHEET"
