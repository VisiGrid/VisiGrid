#!/usr/bin/env bash
# Custom Functions Demo: Verifiable User-Defined Computation
#
# This demo shows:
#   1. Build a sheet with custom function formulas
#   2. Fingerprint the content
#   3. Change an input → fingerprint changes
#   4. Revert the input → fingerprint matches
#
# Prerequisites:
#   cp examples/custom-functions/functions.lua ~/.config/visigrid/functions.lua
#
# Run: ./examples/custom-functions/demo.sh

set -euo pipefail

# Find visigrid binary
VISIGRID="${VISIGRID:-visigrid-cli}"
if ! command -v "$VISIGRID" >/dev/null 2>&1; then
    if [ -f "./target/release/visigrid-cli" ]; then
        VISIGRID="./target/release/visigrid-cli"
    elif [ -f "./target/debug/visigrid-cli" ]; then
        VISIGRID="./target/debug/visigrid-cli"
    else
        echo "visigrid-cli not found. Build with: cargo build --release -p visigrid-cli" >&2
        exit 1
    fi
fi

DIR="$(cd "$(dirname "$0")" && pwd)"
OUTPUT="/tmp/visigrid_bond_portfolio.sheet"

echo "═══════════════════════════════════════════════════════════════"
echo "  Custom Functions Demo: Verifiable Computation"
echo "═══════════════════════════════════════════════════════════════"
echo ""

# ─── Step 1: Build ───────────────────────────────────────────────
echo "Step 1: Build sheet from Lua (with custom function formulas)"
echo "──────────────────────────────────────────────────────────────"
echo "$ visigrid-cli sheet apply portfolio.sheet --lua bond_portfolio.lua --stamp --json"
echo ""

RESULT=$("$VISIGRID" sheet apply "$OUTPUT" --lua "$DIR/bond_portfolio.lua" --stamp --json)
echo "$RESULT" | jq .
echo ""

# ─── Step 2: Fingerprint ────────────────────────────────────────
echo "Step 2: Fingerprint (the trust anchor)"
echo "──────────────────────────────────────────────────────────────"
echo "$ visigrid-cli sheet fingerprint portfolio.sheet --json"

FP1_RESULT=$("$VISIGRID" sheet fingerprint "$OUTPUT" --json)
echo "$FP1_RESULT" | jq .
FP1=$(echo "$FP1_RESULT" | jq -r .fingerprint)
echo ""
echo "  Fingerprint: $FP1"
echo ""

# ─── Step 3: Change an input ────────────────────────────────────
echo "Step 3: Change Bond 1 principal (1,000,000 → 2,000,000)"
echo "──────────────────────────────────────────────────────────────"

# Apply a one-liner that changes just the principal
"$VISIGRID" sheet apply "$OUTPUT" --lua <(echo 'set("B6", 2000000)') --json > /dev/null

echo "$ visigrid-cli sheet fingerprint portfolio.sheet --json"
FP2_RESULT=$("$VISIGRID" sheet fingerprint "$OUTPUT" --json)
echo "$FP2_RESULT" | jq .
FP2=$(echo "$FP2_RESULT" | jq -r .fingerprint)
echo ""

if [ "$FP1" != "$FP2" ]; then
    echo "  Fingerprint CHANGED: $FP1 → $FP2"
    echo "  (Input changed → fingerprint changed. Drift detected.)"
else
    echo "  ERROR: fingerprint should have changed!"
    exit 1
fi
echo ""

# ─── Step 4: Revert the input ───────────────────────────────────
echo "Step 4: Revert Bond 1 principal (2,000,000 → 1,000,000)"
echo "──────────────────────────────────────────────────────────────"

"$VISIGRID" sheet apply "$OUTPUT" --lua <(echo 'set("B6", 1000000)') --json > /dev/null

echo "$ visigrid-cli sheet fingerprint portfolio.sheet --json"
FP3_RESULT=$("$VISIGRID" sheet fingerprint "$OUTPUT" --json)
echo "$FP3_RESULT" | jq .
FP3=$(echo "$FP3_RESULT" | jq -r .fingerprint)
echo ""

if [ "$FP1" = "$FP3" ]; then
    echo "  Fingerprint MATCHES original: $FP3"
    echo "  (Same inputs + same formulas = same fingerprint. Always.)"
else
    echo "  ERROR: fingerprint should match original!"
    echo "  Original: $FP1"
    echo "  Current:  $FP3"
    exit 1
fi
echo ""

# ─── Step 5: Verify ─────────────────────────────────────────────
echo "Step 5: Verify against stamped fingerprint"
echo "──────────────────────────────────────────────────────────────"

# Re-stamp with original build
"$VISIGRID" sheet apply "$OUTPUT" --lua "$DIR/bond_portfolio.lua" --stamp --json > /dev/null
echo "$ visigrid-cli sheet verify portfolio.sheet --fingerprint $FP1"
"$VISIGRID" sheet verify "$OUTPUT" --fingerprint "$FP1"
echo ""

# ─── Summary ─────────────────────────────────────────────────────
echo "═══════════════════════════════════════════════════════════════"
echo "  DEMO COMPLETE"
echo ""
echo "  What you just saw:"
echo "    1. Built a bond portfolio with =ACCRUED_INTEREST() formulas"
echo "    2. Fingerprinted it (inputs + formulas, not style)"
echo "    3. Changed an input → fingerprint changed"
echo "    4. Reverted the input → fingerprint matched"
echo "    5. Verified against the stamped fingerprint"
echo ""
echo "  Why this matters:"
echo "    Excel and Sheets cannot do this."
echo "    Same script + same engine = same output. Provably."
echo ""
echo "  Next: Open in VisiGrid to see the custom functions evaluate:"
echo "    visigrid $OUTPUT"
echo "═══════════════════════════════════════════════════════════════"

# Cleanup (optional)
# rm -f "$OUTPUT"
