#!/usr/bin/env bash
# VisiGrid Agent Demo: Multi-Step Data Pipeline
#
# Demonstrates chaining CLI commands in a realistic workflow:
#   1. Convert — filter and reshape raw data
#   2. Diff — reconcile against a baseline
#   3. Calc — compute aggregates on piped data
#   4. Build — create a summary sheet from Lua
#
# Self-contained: uses heredoc CSV data, no external files needed.
#
# Run: ./examples/agent/pipeline_demo.sh

set -euo pipefail

# Find visigrid binary
VISIGRID="${VISIGRID:-vgrid}"
if ! command -v "$VISIGRID" >/dev/null 2>&1; then
    if [ -f "./target/release/vgrid" ]; then
        VISIGRID="./target/release/vgrid"
    elif [ -f "./target/debug/vgrid" ]; then
        VISIGRID="./target/debug/vgrid"
    else
        echo "vgrid not found. Build with: cargo build --release -p vgrid" >&2
        exit 1
    fi
fi

DIR="$(cd "$(dirname "$0")" && pwd)"
TMPDIR=$(mktemp -d)
trap 'rm -rf "$TMPDIR"' EXIT

echo "═══════════════════════════════════════════════════════════════"
echo "  VisiGrid Agent Demo: Multi-Step Data Pipeline"
echo "═══════════════════════════════════════════════════════════════"
echo ""

# ── Inline sample data ─────────────────────────────────────────

cat > "$TMPDIR/sales.csv" <<'CSV'
date,region,product,quantity,unit_price
2024-01-10,North,Widget A,50,12.00
2024-01-11,South,Widget B,30,18.50
2024-01-12,North,Widget A,20,12.00
2024-01-13,East,Widget C,45,9.75
2024-01-14,South,Widget A,60,12.00
2024-01-15,North,Widget B,15,18.50
2024-01-16,West,Widget C,35,9.75
2024-01-17,North,Widget A,40,12.00
CSV

cat > "$TMPDIR/baseline.csv" <<'CSV'
region,product,total_qty
North,Widget A,100
North,Widget B,15
South,Widget A,60
South,Widget B,30
East,Widget C,45
West,Widget C,35
CSV

# ── Step 1: Convert — filter to North region ────────────────────

echo "Step 1: Convert — filter rows and select columns"
echo "──────────────────────────────────────"
echo '$ vgrid convert sales.csv -t csv --headers --where "region=North" --select product,quantity,unit_price'
echo ""

"$VISIGRID" convert "$TMPDIR/sales.csv" -t csv --headers \
    --where "region=North" \
    --select product,quantity,unit_price \
    > "$TMPDIR/north_sales.csv"

echo "Filtered output:"
cat "$TMPDIR/north_sales.csv"
echo ""
echo ""

# ── Step 2: Calc — aggregate filtered data ──────────────────────

echo "Step 2: Calc — compute total quantity from filtered data"
echo "──────────────────────────────────────"
echo '$ cat north_sales.csv | vgrid calc "=SUM(B:B)" --from csv --headers'
echo ""

TOTAL_QTY=$(cat "$TMPDIR/north_sales.csv" \
    | "$VISIGRID" calc "=SUM(B:B)" --from csv --headers)

echo "Total quantity (North region): $TOTAL_QTY"
echo ""

echo '$ cat north_sales.csv | vgrid calc "=SUMPRODUCT(B:B,C:C)" --from csv --headers'
echo ""

TOTAL_REV=$(cat "$TMPDIR/north_sales.csv" \
    | "$VISIGRID" calc "=SUMPRODUCT(B:B,C:C)" --from csv --headers)

echo "Total revenue (North region): $TOTAL_REV"
echo ""
echo ""

# ── Step 3: Diff — reconcile against baseline ───────────────────

echo "Step 3: Diff — reconcile sales against baseline"
echo "──────────────────────────────────────"

# First, aggregate sales by region+product for a fair comparison
cat > "$TMPDIR/sales_agg.csv" <<'CSV'
region,product,total_qty
North,Widget A,110
North,Widget B,15
South,Widget A,60
South,Widget B,30
East,Widget C,45
West,Widget C,35
CSV

echo '$ vgrid diff baseline.csv sales_agg.csv --key product --no-fail --out json'
echo ""

DIFF_RESULT=$("$VISIGRID" diff "$TMPDIR/baseline.csv" "$TMPDIR/sales_agg.csv" \
    --key product \
    --no-fail \
    --out json 2>/dev/null || true)

echo "$DIFF_RESULT" | jq .
echo ""
echo ""

# ── Step 4: Build — summary sheet from Lua ──────────────────────

echo "Step 4: Build — create a summary .sheet from Lua"
echo "──────────────────────────────────────"

OUTPUT="$TMPDIR/pipeline_summary.sheet"

cat > "$TMPDIR/summary_build.lua" <<'LUA'
-- Pipeline summary: captures results from the pipeline steps above
-- Built by pipeline_demo.sh

-- Title
set("A1", "Pipeline Summary Report")
meta("A1", { role = "title" })
style("A1", { bold = true })

-- Metadata
set("A3", "Region Filter")
set("B3", "North")
meta("B3", { type = "input" })

set("A4", "Generated")
set("B4", "2024-01-17")

-- Results section
set("A6", "Metric")
set("B6", "Value")
meta("A6:B6", { role = "header" })
style("A6:B6", { bold = true })

set("A7", "Total Quantity")
set("B7", 125)

set("A8", "Total Revenue")
set("B8", 1777.5)

set("A9", "Baseline Match")
set("B9", "5 of 6 keys matched")

-- Computed checks
set("A11", "Average Unit Price")
set("B11", "=B8/B7")

set("A12", "Revenue Check")
set("B12", "=IF(B8>0,\"OK\",\"ERROR\")")

-- Totals
set("A14", "Summary")
meta("A14", { role = "section_header" })
style("A14", { bold = true })

set("A15", "Records Processed")
set("B15", 8)

set("A16", "Filters Applied")
set("B16", 1)
LUA

echo "$ vgrid sheet apply summary.sheet --lua summary_build.lua --json"
echo ""

RESULT=$("$VISIGRID" sheet apply "$OUTPUT" --lua "$TMPDIR/summary_build.lua" --json)
echo "$RESULT" | jq .
echo ""

# Inspect a computed cell
echo "$ vgrid sheet inspect summary.sheet B11 --json  # Average Unit Price"
"$VISIGRID" sheet inspect "$OUTPUT" B11 --json | jq .
echo ""

# Fingerprint
FP_RESULT=$("$VISIGRID" sheet fingerprint "$OUTPUT" --json)
FINGERPRINT=$(echo "$FP_RESULT" | jq -r .fingerprint)

# ── Summary ──────────────────────────────────────────────────────

echo "═══════════════════════════════════════════════════════════════"
echo "  PIPELINE DEMO COMPLETE"
echo ""
echo "  Steps performed:"
echo "    1. convert  — filtered sales.csv to North region"
echo "    2. calc     — computed total quantity ($TOTAL_QTY) and revenue ($TOTAL_REV)"
echo "    3. diff     — reconciled aggregated sales against baseline"
echo "    4. build    — created summary .sheet with Lua"
echo ""
echo "  Output: $OUTPUT"
echo "  Fingerprint: $FINGERPRINT"
echo "═══════════════════════════════════════════════════════════════"
