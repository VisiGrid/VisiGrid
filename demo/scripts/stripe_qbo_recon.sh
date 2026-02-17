#!/usr/bin/env bash
# stripe_qbo_recon.sh — Stripe ↔ QBO reconciliation demo
#
# Pipeline: fetch both sources → fill recon template → inspect summary → publish
#
# The recon template (stripe-qbo-recon.sheet) is the reconciliation engine:
#   - XLOOKUP matches Stripe payouts to QBO deposits by amount
#   - Summary sheet computes totals, match counts, and pass/fail status
#   - vgrid computes everything; this script only orchestrates + reads results
#
# Usage:
#   # Fixture data (no API keys needed)
#   ./demo/scripts/stripe_qbo_recon.sh --mode pass --no-publish
#   ./demo/scripts/stripe_qbo_recon.sh --mode fail --no-publish
#
#   # Publish fixture results to VisiHub
#   ./demo/scripts/stripe_qbo_recon.sh --repo ORG/REPO --mode pass
#   ./demo/scripts/stripe_qbo_recon.sh --repo ORG/REPO --mode fail
#
#   # Live API data
#   ./demo/scripts/stripe_qbo_recon.sh --repo ORG/REPO --mode live \
#     --qbo-credentials ~/.config/vgrid/qbo.json --qbo-account "Checking"
#
# Prerequisites:
#   - vgrid on PATH
#   - For publish: vgrid login
#   - For live mode: STRIPE_API_KEY set (or --stripe-api-key)
#                    QBO credentials file (--qbo-credentials)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEMO_DIR="$(dirname "$SCRIPT_DIR")"
TEMPLATE="$DEMO_DIR/templates/stripe-qbo-recon.sheet"
FIXTURE_DIR="$DEMO_DIR/data/stripe-qbo"

# ── Defaults ───────────────────────────────────────────────────────
FROM="$(date -d '45 days ago' +%Y-%m-%d 2>/dev/null || date -v-45d +%Y-%m-%d)"
TO="$(date +%Y-%m-%d)"
REPO=""
DATASET="stripe-qbo-recon"
MODE="pass"
STRIPE_API_KEY=""
QBO_CREDENTIALS=""
QBO_ACCOUNT=""
QBO_ACCOUNT_ID=""
QBO_SANDBOX=""
PUBLISH="true"

# ── Parse flags ────────────────────────────────────────────────────
while [[ $# -gt 0 ]]; do
    case "$1" in
        --from)              FROM="$2";              shift 2 ;;
        --to)                TO="$2";                shift 2 ;;
        --repo)              REPO="$2";              shift 2 ;;
        --dataset)           DATASET="$2";           shift 2 ;;
        --mode)              MODE="$2";              shift 2 ;;
        --stripe-api-key)    STRIPE_API_KEY="$2";    shift 2 ;;
        --qbo-credentials)   QBO_CREDENTIALS="$2";   shift 2 ;;
        --qbo-account)       QBO_ACCOUNT="$2";       shift 2 ;;
        --qbo-account-id)    QBO_ACCOUNT_ID="$2";    shift 2 ;;
        --qbo-sandbox)       QBO_SANDBOX="true";     shift 1 ;;
        --no-publish)        PUBLISH="false";        shift 1 ;;
        *) echo "error: unknown flag $1" >&2; exit 2 ;;
    esac
done

case "$MODE" in
    pass|fail|live) ;;
    *) echo "error: --mode must be pass, fail, or live" >&2; exit 2 ;;
esac

if [[ "$PUBLISH" == "true" && -z "$REPO" ]]; then
    echo "error: --repo required (or use --no-publish)" >&2; exit 2
fi

if [[ "$MODE" == "live" && -z "$QBO_CREDENTIALS" ]]; then
    echo "error: --qbo-credentials required for live mode" >&2; exit 2
fi

if [[ "$MODE" == "live" && -z "$QBO_ACCOUNT" && -z "$QBO_ACCOUNT_ID" ]]; then
    echo "error: --qbo-account or --qbo-account-id required for live mode" >&2; exit 2
fi

# ── Temp files ─────────────────────────────────────────────────────
STRIPE_CSV=$(mktemp --suffix=.csv)
QBO_CSV=$(mktemp --suffix=.csv)
FILLED_SHEET=$(mktemp --suffix=.sheet)
INTERMEDIATE_SHEET=$(mktemp --suffix=.sheet)
PROOF_JSON=""
trap 'rm -f "$STRIPE_CSV" "$QBO_CSV" "$FILLED_SHEET" "$INTERMEDIATE_SHEET" "$PROOF_JSON"' EXIT

echo "=== Stripe ↔ QBO Reconciliation ($MODE) ==="
echo "Window:  $FROM to $TO"
if [[ -n "$REPO" ]]; then echo "Repo:    $REPO"; fi
echo "Dataset: $DATASET"
echo ""

# ── Step 1: Fetch Stripe ──────────────────────────────────────────

if [[ "$MODE" == "live" ]]; then
    echo "--- Fetch Stripe transactions ---"
    FETCH_ARGS=(fetch stripe --from "$FROM" --to "$TO" --out "$STRIPE_CSV")
    [[ -n "$STRIPE_API_KEY" ]] && FETCH_ARGS+=(--api-key "$STRIPE_API_KEY")
    vgrid "${FETCH_ARGS[@]}"
else
    echo "--- Using Stripe fixture data ---"
    cp "$FIXTURE_DIR/stripe.csv" "$STRIPE_CSV"
fi

STRIPE_ROWS=$(tail -n +2 "$STRIPE_CSV" | grep -c . || true)
echo "  $STRIPE_ROWS Stripe transactions"
echo ""

# ── Step 2: Fetch QBO ─────────────────────────────────────────────

if [[ "$MODE" == "live" ]]; then
    echo "--- Fetch QBO posted transactions ---"
    FETCH_ARGS=(fetch qbo --from "$FROM" --to "$TO" --out "$QBO_CSV" --credentials "$QBO_CREDENTIALS")
    [[ -n "$QBO_ACCOUNT" ]] && FETCH_ARGS+=(--account "$QBO_ACCOUNT")
    [[ -n "$QBO_ACCOUNT_ID" ]] && FETCH_ARGS+=(--account-id "$QBO_ACCOUNT_ID")
    [[ -n "$QBO_SANDBOX" ]] && FETCH_ARGS+=(--sandbox)
    vgrid "${FETCH_ARGS[@]}"
elif [[ "$MODE" == "pass" ]]; then
    echo "--- Using QBO fixture data (matching) ---"
    cp "$FIXTURE_DIR/qbo-good.csv" "$QBO_CSV"
else
    echo "--- Using QBO fixture data (mismatched) ---"
    cp "$FIXTURE_DIR/qbo-bad.csv" "$QBO_CSV"
fi

QBO_ROWS=$(tail -n +2 "$QBO_CSV" | grep -c . || true)
echo "  $QBO_ROWS QBO transactions"
echo ""

# ── Step 3: Fill recon template ───────────────────────────────────
#
# Two passes (vgrid fill takes one CSV at a time):
#   Pass 1: Stripe data → stripe sheet
#   Pass 2: QBO data → qbo sheet (from pass 1 output)
#
# The template's XLOOKUP formulas automatically match Stripe payouts
# to QBO deposits by amount. The summary sheet computes totals and
# reports pass/fail.

echo "--- Fill recon template ---"

vgrid fill "$TEMPLATE" \
    --csv "$STRIPE_CSV" \
    --target "stripe!A2" --headers --clear \
    --out "$INTERMEDIATE_SHEET" --json
echo ""

vgrid fill "$INTERMEDIATE_SHEET" \
    --csv "$QBO_CSV" \
    --target "qbo!A2" --headers --clear \
    --out "$FILLED_SHEET" --json
echo ""

# ── Step 4: Read summary (all computed by vgrid) ──────────────────

echo "--- Reconciliation Summary ---"
echo ""

# Helper: read a cell value from the filled workbook
read_cell() {
    vgrid sheet inspect "$FILLED_SHEET" --sheet summary "$1" --json \
        | python3 -c "import sys,json; print(json.load(sys.stdin)['value'])"
}

# Stripe internal balance
B2=$(read_cell B2)   # Charges
B3=$(read_cell B3)   # Fees
B4=$(read_cell B4)   # Refunds
B5=$(read_cell B5)   # Payouts
B6=$(read_cell B6)   # Adjustments
B7=$(read_cell B7)   # Net
C7=$(read_cell C7)   # Pass/Fail

echo "  STRIPE BALANCE"
echo "    Charges:       $B2"
echo "    Fees:          $B3"
echo "    Refunds:       $B4"
echo "    Payouts:       $B5"
echo "    Adjustments:   $B6"
echo "    Net:           $B7  [$C7]"
echo ""

# Payout matching (the core reconciliation)
B10=$(read_cell B10)  # Stripe Payouts (abs)
B11=$(read_cell B11)  # Matched Deposits
B12=$(read_cell B12)  # Unmatched Payouts
B13=$(read_cell B13)  # Difference
C13=$(read_cell C13)  # Pass/Fail

echo "  PAYOUT MATCHING"
echo "    Stripe Payouts:    $B10"
echo "    Matched Deposits:  $B11"
echo "    Unmatched Payouts: $B12"
echo "    Difference:        $B13  [$C13]"
echo ""

# Match counts
B16=$(read_cell B16)  # Stripe Payouts (count)
B17=$(read_cell B17)  # Matched (count)
B18=$(read_cell B18)  # Unmatched (count)

echo "  MATCH COUNTS"
echo "    Stripe Payouts:  $B16"
echo "    Matched:         $B17"
echo "    Unmatched:       $B18"
echo ""

# QBO unmatched
B21=$(read_cell B21)  # Total Deposits (count)
C21=$(read_cell C21)  # Total Deposits (amount)
B22=$(read_cell B22)  # Matched to Stripe (count)
B23=$(read_cell B23)  # Unmatched (count)
C23=$(read_cell C23)  # Unmatched (amount)

echo "  QBO DEPOSITS"
echo "    Total:           $B21 deposits ($C21 minor units)"
echo "    Matched:         $B22"
echo "    Unmatched:       $B23 ($C23 minor units)"
echo ""

# ── Step 5: Sign certificate ──────────────────────────────────────

echo "--- Sign certificate ---"
PROOF_JSON=$(mktemp --suffix=.json)
vgrid sign "$FILLED_SHEET" --out "$PROOF_JSON" --quiet
KEY_ID=$(python3 -c "import sys,json; print(json.load(sys.stdin)['key_id'])" < "$PROOF_JSON")
echo "  Key ID: $KEY_ID"
echo "  Proof:  $PROOF_JSON"
echo ""

# ── Step 6: Publish (optional) ────────────────────────────────────

if [[ "$PUBLISH" == "true" ]]; then
    echo "--- Publish filled workbook ---"
    PUBLISH_ARGS=(publish "$FILLED_SHEET"
        --repo "$REPO"
        --dataset "$DATASET"
        --source-type stripe
        --wait
        --output json)

    # Assert payout matching passes (summary!B13 = 0)
    if [[ "$MODE" != "fail" ]]; then
        PUBLISH_ARGS+=(--assert-cell "summary!B13:0:0")
    fi

    vgrid "${PUBLISH_ARGS[@]}" || true
    echo ""
fi

# ── Result ────────────────────────────────────────────────────────

echo "=== Done ==="
if [[ "$C13" == "PASS" ]]; then
    echo "Result: RECONCILED"
    echo "  Processor payouts match bank deposits for this period."
else
    echo "Result: DIFFERENCES FOUND"
    echo "  $B18 unmatched Stripe payout(s), $B23 unmatched QBO deposit(s)."
    echo "  Variance: $B13 minor units."
fi
if [[ -n "$REPO" ]]; then
    echo "View results: https://visihub.app/$REPO"
fi

exit 0
