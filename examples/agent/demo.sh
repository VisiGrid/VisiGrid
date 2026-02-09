#!/usr/bin/env bash
# VisiGrid Agent Demo: Verifiable Spreadsheet Build
#
# This demo shows the complete agent workflow:
#   1. Build a .sheet from Lua (replacement semantics)
#   2. Inspect the results
#   3. Verify the fingerprint
#
# Run: ./examples/agent/demo.sh

set -euo pipefail

# Find visigrid binary
VISIGRID="${VISIGRID:-vgrid}"
if ! command -v "$VISIGRID" >/dev/null 2>&1; then
    # Try common locations
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
OUTPUT="/tmp/visigrid_demo_model.sheet"

echo "═══════════════════════════════════════════════════════════════"
echo "  VisiGrid Agent Demo: Verifiable Spreadsheet Build"
echo "═══════════════════════════════════════════════════════════════"
echo ""

# Step 1: Build
echo "Step 1: Build .sheet from Lua script"
echo "──────────────────────────────────────"
echo "$ vgrid sheet apply model.sheet --lua revenue_model.lua --json"
echo ""

RESULT=$("$VISIGRID" sheet apply "$OUTPUT" --lua "$DIR/revenue_model.lua" --json)
echo "$RESULT" | jq .
echo ""

# Note: apply fingerprint includes meta() ops; file fingerprint is content-only.
# For verification, we use the file fingerprint (what's actually stored).
BUILD_FP=$(echo "$RESULT" | jq -r .fingerprint)
echo "Build fingerprint (includes meta): $BUILD_FP"
echo ""

# Step 2: Inspect
echo "Step 2: Inspect key cells"
echo "──────────────────────────────────────"

echo "$ vgrid sheet inspect model.sheet B4 --json  # Base Revenue (input)"
"$VISIGRID" sheet inspect "$OUTPUT" B4 --json | jq .
echo ""

echo "$ vgrid sheet inspect model.sheet B19 --json  # Month 12 (formula)"
"$VISIGRID" sheet inspect "$OUTPUT" B19 --json | jq .
echo ""

echo "$ vgrid sheet inspect model.sheet B21 --json  # Total (formula)"
"$VISIGRID" sheet inspect "$OUTPUT" B21 --json | jq .
echo ""

# Step 3: Fingerprint
echo "Step 3: Get file fingerprint (content verification)"
echo "──────────────────────────────────────"
echo "$ vgrid sheet fingerprint model.sheet --json"
FP_RESULT=$("$VISIGRID" sheet fingerprint "$OUTPUT" --json)
echo "$FP_RESULT" | jq .
FINGERPRINT=$(echo "$FP_RESULT" | jq -r .fingerprint)
echo ""

# Step 4: Verify
echo "Step 4: Verify fingerprint (the trust proof)"
echo "──────────────────────────────────────"
echo "$ vgrid sheet verify model.sheet --fingerprint $FINGERPRINT"
"$VISIGRID" sheet verify "$OUTPUT" --fingerprint "$FINGERPRINT"
echo ""

# Summary
echo "═══════════════════════════════════════════════════════════════"
echo "  DEMO COMPLETE"
echo ""
echo "  Workflow: Lua script → apply → inspect → verify"
echo ""
echo "  The fingerprint proves:"
echo "    - Same script + same engine = same output"
echo "    - No hidden state, no ambient context"
echo "    - Agents can build, humans can verify"
echo ""
echo "  Output: $OUTPUT"
echo "  Fingerprint: $FINGERPRINT"
echo "═══════════════════════════════════════════════════════════════"

# Cleanup (optional - comment out to keep the file)
# rm -f "$OUTPUT"

echo ""
echo "  For multi-sheet workbooks, see: multi_sheet_model.lua"
echo "  Run: vgrid sheet apply workbook.sheet --lua multi_sheet_model.lua --json"
