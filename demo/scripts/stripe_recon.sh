#!/usr/bin/env bash
# stripe_recon.sh — Fetch Stripe → fill recon template → publish signed proof
#
# Usage:
#   ./demo/scripts/stripe_recon.sh --repo ORG/REPO
#   ./demo/scripts/stripe_recon.sh --repo ORG/REPO --from 2026-01-01 --to 2026-02-01
#   ./demo/scripts/stripe_recon.sh --repo ORG/REPO --mode break
#
# Prerequisites:
#   - vgrid on PATH
#   - STRIPE_API_KEY set (or pass --api-key)
#   - Authenticated: vgrid login

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEMO_DIR="$(dirname "$SCRIPT_DIR")"
TEMPLATE="$DEMO_DIR/templates/recon-template.sheet"

# ── Defaults ───────────────────────────────────────────────────────
FROM="$(date -d '45 days ago' +%Y-%m-%d)"
TO="$(date +%Y-%m-%d)"
REPO=""
DATASET="ledger-recon"
API_KEY=""
TOLERANCE="0"
MODE="baseline"

# ── Parse flags ────────────────────────────────────────────────────
while [[ $# -gt 0 ]]; do
    case "$1" in
        --from)      FROM="$2";      shift 2 ;;
        --to)        TO="$2";        shift 2 ;;
        --repo)      REPO="$2";      shift 2 ;;
        --dataset)   DATASET="$2";   shift 2 ;;
        --api-key)   API_KEY="$2";   shift 2 ;;
        --tolerance) TOLERANCE="$2"; shift 2 ;;
        --mode)      MODE="$2";      shift 2 ;;
        *) echo "error: unknown flag $1" >&2; exit 2 ;;
    esac
done

[[ -z "$REPO" ]] && { echo "error: --repo required" >&2; exit 2; }

case "$MODE" in
    baseline|assert-zero|break) ;;
    *) echo "error: --mode must be baseline, assert-zero, or break" >&2; exit 2 ;;
esac

STRIPE_CSV=$(mktemp --suffix=.csv)
RECON_SHEET=$(mktemp --suffix=.sheet)
trap 'rm -f "$STRIPE_CSV" "$RECON_SHEET"' EXIT

echo "=== Stripe Reconciliation ($MODE) ==="
echo "Window:  $FROM to $TO"
echo "Repo:    $REPO"
echo "Dataset: $DATASET"
echo ""

# ── Step 1: Fetch ──────────────────────────────────────────────────

if [[ "$MODE" == "break" ]]; then
    echo "--- Using fixture data (intentional invariant violation) ---"
    cp "$DEMO_DIR/data/ledger-bad-undistributed.csv" "$STRIPE_CSV"
else
    echo "--- Fetch Stripe transactions ---"
    FETCH_ARGS=(fetch stripe --from "$FROM" --to "$TO" --out "$STRIPE_CSV")
    [[ -n "$API_KEY" ]] && FETCH_ARGS+=(--api-key "$API_KEY")
    vgrid "${FETCH_ARGS[@]}"
fi

ROW_COUNT=$(tail -n +2 "$STRIPE_CSV" | grep -c . || true)
echo "  $ROW_COUNT transactions"
echo ""

# ── Step 2: Fill template ──────────────────────────────────────────

echo "--- Fill recon template ---"
vgrid fill "$TEMPLATE" \
    --csv "$STRIPE_CSV" \
    --target "tx!A2" --headers --clear \
    --out "$RECON_SHEET" --json
echo ""

# ── Read B7 ────────────────────────────────────────────────────────

B7=$(vgrid sheet inspect "$RECON_SHEET" --sheet summary B7 --json \
    | grep -o '"display":"[^"]*"' | head -1 | sed 's/"display":"//;s/"$//')

echo "Computed undistributed balance: $B7 (minor units)"
echo ""

# ── Step 3: Publish ────────────────────────────────────────────────

case "$MODE" in
    baseline)
        echo "--- Publish (baseline — no assertion) ---"
        vgrid publish "$RECON_SHEET" \
            --repo "$REPO" \
            --dataset "$DATASET" \
            --source-type stripe \
            --wait \
            --output json
        ;;
    assert-zero)
        echo "--- Publish with assertion (summary!B7 = 0 ± $TOLERANCE) ---"
        vgrid publish "$RECON_SHEET" \
            --repo "$REPO" \
            --dataset "$DATASET" \
            --source-type stripe \
            --wait \
            --assert-cell "summary!B7:0:$TOLERANCE" \
            --output json
        ;;
    break)
        echo "--- Publish with assertion (expecting failure) ---"
        vgrid publish "$RECON_SHEET" \
            --repo "$REPO" \
            --dataset "$DATASET" \
            --source-type stripe \
            --wait \
            --assert-cell "summary!B7:0:0" \
            --output json || true
        ;;
esac

echo ""
echo "=== Done ==="
echo "View results: https://visihub.app/$REPO"
