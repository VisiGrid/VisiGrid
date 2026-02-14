#!/usr/bin/env bash
# run_local.sh — Local smoke test for VisiGrid recon pipeline
#
# Prerequisites:
#   - vgrid CLI built and on PATH (cargo install --path crates/cli)
#   - VISIHUB_REPO set (e.g., robert/test-repo)
#   - Authenticated: vgrid login
#
# Usage:
#   VISIHUB_REPO=robert/test-repo ./demo/scripts/run_local.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEMO_DIR="$(dirname "$SCRIPT_DIR")"
REPO="${VISIHUB_REPO:?Set VISIHUB_REPO (e.g., robert/test-repo)}"
DATASET="recon-demo-$(date +%s)"
TEMPLATE="$DEMO_DIR/templates/recon-template.sheet"
GOOD_CSV="$DEMO_DIR/data/ledger-good.csv"
BAD_APPEND="$DEMO_DIR/data/ledger-bad-append.csv"
BAD_REMOVE="$DEMO_DIR/data/ledger-bad-remove.csv"
BAD_COLUMN="$DEMO_DIR/data/ledger-bad-column-add.csv"

echo "=== VisiGrid Recon Demo ==="
echo "Repo:    $REPO"
echo "Dataset: $DATASET"
echo ""

# Step 1: Fill template with good data
FILLED=$(mktemp --suffix=.sheet)
echo "--- Fill template ---"
vgrid fill "$TEMPLATE" \
  --csv "$GOOD_CSV" \
  --target "tx!A2" \
  --headers --clear \
  --out "$FILLED" --json
echo ""

# Step 2: Publish baseline (CSV)
echo "--- Publish baseline ---"
vgrid publish "$GOOD_CSV" \
  --repo "$REPO" \
  --dataset "$DATASET" \
  --source-type manual \
  --wait --output json
echo ""

# Step 3: Repeat identical
echo "--- Repeat identical ---"
vgrid publish "$GOOD_CSV" \
  --repo "$REPO" \
  --dataset "$DATASET" \
  --source-type manual \
  --wait --output json
echo ""

# Step 4: Append rows (expect fail/warn)
echo "--- Append rows ---"
vgrid publish "$BAD_APPEND" \
  --repo "$REPO" \
  --dataset "$DATASET" \
  --source-type manual \
  --wait --no-fail --output json
echo ""

# Step 5: Remove rows (expect fail)
echo "--- Remove rows ---"
vgrid publish "$BAD_REMOVE" \
  --repo "$REPO" \
  --dataset "$DATASET" \
  --source-type manual \
  --wait --no-fail --output json
echo ""

# Step 6: Add column (expect fail/warn)
echo "--- Add column ---"
vgrid publish "$BAD_COLUMN" \
  --repo "$REPO" \
  --dataset "$DATASET" \
  --source-type manual \
  --wait --no-fail --output json
echo ""

# Cleanup
rm -f "$FILLED"

echo "=== Done ==="
echo "View results: https://visihub.app/$REPO"
