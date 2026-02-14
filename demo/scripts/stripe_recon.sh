#!/usr/bin/env bash
# stripe_recon.sh — Fetch from Stripe, compile into a deterministic sheet,
#                    publish a signed proof. If the invariant breaks, CI fails.
#
# Usage:
#   ./demo/scripts/stripe_recon.sh \
#     --from 2026-02-01 --to 2026-02-02 \
#     --repo ORG/REPO --dataset ledger-recon
#
# Prerequisites:
#   - vgrid on PATH
#   - STRIPE_API_KEY set (or pass --api-key)
#   - Authenticated: vgrid login

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEMO_DIR="$(dirname "$SCRIPT_DIR")"
TEMPLATE="$DEMO_DIR/templates/recon-template.sheet"

# ── Parse flags ──────────────────────────────────────────────────────

FROM=""
TO=""
REPO=""
DATASET="ledger-recon"
API_KEY=""
TOLERANCE="0"

while [[ $# -gt 0 ]]; do
    case "$1" in
        --from)      FROM="$2";      shift 2 ;;
        --to)        TO="$2";        shift 2 ;;
        --repo)      REPO="$2";      shift 2 ;;
        --dataset)   DATASET="$2";   shift 2 ;;
        --api-key)   API_KEY="$2";   shift 2 ;;
        --tolerance) TOLERANCE="$2"; shift 2 ;;
        *) echo "error: unknown flag $1" >&2; exit 2 ;;
    esac
done

[[ -z "$FROM" ]] && { echo "error: --from required" >&2; exit 2; }
[[ -z "$TO" ]]   && { echo "error: --to required" >&2; exit 2; }
[[ -z "$REPO" ]] && { echo "error: --repo required" >&2; exit 2; }

STRIPE_CSV=$(mktemp --suffix=.csv)
RECON_SHEET=$(mktemp --suffix=.sheet)
trap 'rm -f "$STRIPE_CSV" "$RECON_SHEET"' EXIT

# ── Step 1: Fetch ────────────────────────────────────────────────────

echo "--- Fetch Stripe transactions ($FROM to $TO) ---"
FETCH_ARGS=(fetch stripe --from "$FROM" --to "$TO" --out "$STRIPE_CSV")
[[ -n "$API_KEY" ]] && FETCH_ARGS+=(--api-key "$API_KEY")
vgrid "${FETCH_ARGS[@]}"

ROW_COUNT=$(tail -n +2 "$STRIPE_CSV" | wc -l)
echo "  $ROW_COUNT transactions"

# ── Step 2: Fill template ────────────────────────────────────────────

echo "--- Fill recon template ---"
vgrid fill "$TEMPLATE" \
    --csv "$STRIPE_CSV" \
    --target "tx!A2" --headers --clear \
    --out "$RECON_SHEET" --json
echo ""

# ── Step 3: Publish with invariant assertion ─────────────────────────

echo "--- Publish with assertion (summary!B7 = 0 ± $TOLERANCE) ---"
vgrid publish "$RECON_SHEET" \
    --repo "$REPO" \
    --dataset "$DATASET" \
    --source-type stripe \
    --wait \
    --assert-cell "summary!B7:0:$TOLERANCE" \
    --output json
echo ""

echo "=== Done ==="
echo "View results: https://visihub.app/$REPO"
