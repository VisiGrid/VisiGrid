#!/usr/bin/env bash
# seed_nakatomi.sh — Publish Nakatomi Corporation demo data via trust pipeline
#
# Fills real .sheet templates with fixture CSV data and publishes to VisiHub
# as the nakatomi user. Creates the PASS → DRIFT → RESOLVED revision story
# that powers the /explore gallery.
#
# Prerequisites:
#   - vgrid on PATH (cargo install --path crates/cli)
#   - Authenticated as nakatomi: vgrid login
#   - Nakatomi user + repos already created on server (rake examples:seed_nakatomi_repos)
#
# Usage:
#   ./demo/scripts/seed_nakatomi.sh
#   ./demo/scripts/seed_nakatomi.sh --dry-run    # Local only, no upload
#   ./demo/scripts/seed_nakatomi.sh --repo 1     # Only seed repo 1 (stripe-qbo)
#   ./demo/scripts/seed_nakatomi.sh --repo 3     # Only seed repo 3 (stripe-brex)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEMO_DIR="$(dirname "$SCRIPT_DIR")"
APP_DIR="$(dirname "$DEMO_DIR")"
TEMPLATE_DIR="$DEMO_DIR/templates"
FIXTURE_DIR="$DEMO_DIR/data"
RECON_CSV_DIR="$APP_DIR/crates/cli/tests/recon/csv"

DRY_RUN="false"
ONLY_REPO=""

while [[ $# -gt 0 ]]; do
    case "$1" in
        --dry-run)  DRY_RUN="true"; shift ;;
        --repo)     ONLY_REPO="$2"; shift 2 ;;
        *) echo "error: unknown flag $1" >&2; exit 2 ;;
    esac
done

PUBLISH_FLAGS=(--json)
if [[ "$DRY_RUN" == "true" ]]; then
    PUBLISH_FLAGS+=(--dry-run)
fi

# Helper: read a cell value from a .sheet file
read_cell() {
    local sheet_file="$1" sheet_name="$2" cell="$3"
    vgrid sheet inspect "$sheet_file" --sheet "$sheet_name" "$cell" --json \
        | python3 -c "import sys,json; print(json.load(sys.stdin)['value'])"
}

# Helper: publish a .sheet with message and optional summary JSON
hub_publish() {
    local sheet_file="$1" repo="$2" message="$3"
    shift 3
    local extra_args=("$@")

    echo "  Publishing: $message"
    vgrid hub publish "$sheet_file" \
        --repo "$repo" \
        --message "$message" \
        "${PUBLISH_FLAGS[@]}" \
        "${extra_args[@]}" || true
    echo ""
}

echo "=== Nakatomi Corporation — Demo Data Seeding ==="
echo ""
if [[ "$DRY_RUN" == "true" ]]; then
    echo "  MODE: dry-run (no upload)"
else
    echo "  MODE: live publish"
fi
echo ""

# ======================================================================
# REPO 1: Stripe → QBO Payout Matching
# ======================================================================
if [[ -z "$ONLY_REPO" || "$ONLY_REPO" == "1" ]]; then
    echo "━━━ REPO 1: nakatomi/stripe-qbo-reconciliation ━━━"
    echo ""

    TEMPLATE="$TEMPLATE_DIR/stripe-qbo-recon.sheet"
    REPO="nakatomi/stripe-qbo-reconciliation"

    # Rev 1: Baseline — all payouts match (use good QBO data)
    echo "--- Rev 1: Baseline (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)
    trap 'rm -f "$INTER" "$FILLED"' EXIT

    vgrid fill "$TEMPLATE" \
        --csv "$FIXTURE_DIR/stripe-qbo/stripe.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$FIXTURE_DIR/stripe-qbo/qbo-good.csv" \
        --target "qbo!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    echo "  Payout matching: $C13"

    hub_publish "$FILLED" "$REPO" "Baseline — 4 payouts reconciled, all matched"
    rm -f "$INTER" "$FILLED"

    # Rev 2: Drift — QBO deposit short by $50
    echo "--- Rev 2: Drift (FAIL) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$FIXTURE_DIR/stripe-qbo/stripe.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$FIXTURE_DIR/stripe-qbo/qbo-bad.csv" \
        --target "qbo!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    B13=$(read_cell "$FILLED" summary B13)
    echo "  Payout matching: $C13 (difference: $B13)"

    hub_publish "$FILLED" "$REPO" "Drift detected — deposit:502 short by \$50.00"
    rm -f "$INTER" "$FILLED"

    # Rev 3: Resolved — back to good data
    echo "--- Rev 3: Resolved (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$FIXTURE_DIR/stripe-qbo/stripe.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$FIXTURE_DIR/stripe-qbo/qbo-good.csv" \
        --target "qbo!A2" --headers --clear \
        --out "$FILLED" --json

    hub_publish "$FILLED" "$REPO" "Resolved — deposit:502 corrected, all matched"
    rm -f "$INTER" "$FILLED"

    echo ""
fi

# ======================================================================
# REPO 2: Stripe → Mercury Daily Reconciliation
# ======================================================================
if [[ -z "$ONLY_REPO" || "$ONLY_REPO" == "2" ]]; then
    echo "━━━ REPO 2: nakatomi/stripe-mercury-reconciliation ━━━"
    echo ""

    TEMPLATE="$TEMPLATE_DIR/stripe-mercury-recon.sheet"
    REPO="nakatomi/stripe-mercury-reconciliation"

    # Rev 1: Baseline — all balanced
    echo "--- Rev 1: Baseline (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/stripe-balanced.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/mercury-balanced.csv" \
        --target "mercury!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    C40=$(read_cell "$FILLED" summary C40)
    echo "  Payout matching: $C13 | Overall: $C40"

    hub_publish "$FILLED" "$REPO" "Baseline — 4 payouts matched to Mercury deposits"
    rm -f "$INTER" "$FILLED"

    # Rev 2: Drift — Mercury missing one deposit
    echo "--- Rev 2: Drift (FAIL) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/stripe-balanced.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/mercury-missing-one.csv" \
        --target "mercury!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    B13=$(read_cell "$FILLED" summary B13)
    echo "  Payout matching: $C13 (difference: $B13)"

    hub_publish "$FILLED" "$REPO" "Drift — Mercury missing deposit for po_1001"
    rm -f "$INTER" "$FILLED"

    # Rev 3: Resolved
    echo "--- Rev 3: Resolved (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/stripe-balanced.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/mercury-balanced.csv" \
        --target "mercury!A2" --headers --clear \
        --out "$FILLED" --json

    hub_publish "$FILLED" "$REPO" "Resolved — all Mercury deposits matched"
    rm -f "$INTER" "$FILLED"

    echo ""
fi

# ======================================================================
# REPO 3: Stripe → Brex Deposit Matching
# ======================================================================
if [[ -z "$ONLY_REPO" || "$ONLY_REPO" == "3" ]]; then
    echo "━━━ REPO 3: nakatomi/stripe-brex-reconciliation ━━━"
    echo ""

    TEMPLATE="$TEMPLATE_DIR/stripe-brex-recon.sheet"
    REPO="nakatomi/stripe-brex-reconciliation"

    # Rev 1: Baseline — balanced
    echo "--- Rev 1: Baseline (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/stripe-balanced.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/brex-balanced.csv" \
        --target "brex!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    echo "  Payout matching: $C13"

    hub_publish "$FILLED" "$REPO" "Baseline — 2 payouts matched to Brex deposits"
    rm -f "$INTER" "$FILLED"

    # Rev 2: Drift — Brex missing one deposit
    echo "--- Rev 2: Drift (FAIL) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/stripe-balanced.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/brex-missing-one.csv" \
        --target "brex!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    B13=$(read_cell "$FILLED" summary B13)
    echo "  Payout matching: $C13 (difference: $B13)"

    hub_publish "$FILLED" "$REPO" "Drift — Brex missing deposit for po_002"
    rm -f "$INTER" "$FILLED"

    # Rev 3: Resolved
    echo "--- Rev 3: Resolved (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/stripe-balanced.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/brex-balanced.csv" \
        --target "brex!A2" --headers --clear \
        --out "$FILLED" --json

    hub_publish "$FILLED" "$REPO" "Resolved — all Brex deposits matched"
    rm -f "$INTER" "$FILLED"

    echo ""
fi

# ======================================================================
# REPO 4: Gusto → Mercury Payroll Matching
# ======================================================================
if [[ -z "$ONLY_REPO" || "$ONLY_REPO" == "4" ]]; then
    echo "━━━ REPO 4: nakatomi/gusto-mercury-payroll ━━━"
    echo ""

    # Gusto-Mercury uses the stripe-mercury template (same structure:
    # two-source matching via XLOOKUP). We fill the "stripe" sheet with
    # Gusto payroll data and the "mercury" sheet with Mercury withdrawals.
    TEMPLATE="$TEMPLATE_DIR/stripe-mercury-recon.sheet"
    REPO="nakatomi/gusto-mercury-payroll"

    # Rev 1: Baseline — all payroll debits matched
    echo "--- Rev 1: Baseline (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/gusto-two-payrolls.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/mercury-payroll-matched.csv" \
        --target "mercury!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    echo "  Matching: $C13"

    hub_publish "$FILLED" "$REPO" "Baseline — 2 payrolls, all debits matched"
    rm -f "$INTER" "$FILLED"

    # Rev 2: Drift — Mercury missing a payroll withdrawal
    echo "--- Rev 2: Drift (FAIL) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/gusto-two-payrolls.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/mercury-payroll-missing.csv" \
        --target "mercury!A2" --headers --clear \
        --out "$FILLED" --json

    C13=$(read_cell "$FILLED" summary C13)
    B13=$(read_cell "$FILLED" summary B13)
    echo "  Matching: $C13 (difference: $B13)"

    hub_publish "$FILLED" "$REPO" "Drift — Mercury missing Jan 16-31 net pay withdrawal"
    rm -f "$INTER" "$FILLED"

    # Rev 3: Resolved
    echo "--- Rev 3: Resolved (PASS) ---"
    INTER=$(mktemp --suffix=.sheet)
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$RECON_CSV_DIR/gusto-two-payrolls.csv" \
        --target "stripe!A2" --headers --clear \
        --out "$INTER" --json
    vgrid fill "$INTER" \
        --csv "$RECON_CSV_DIR/mercury-payroll-matched.csv" \
        --target "mercury!A2" --headers --clear \
        --out "$FILLED" --json

    hub_publish "$FILLED" "$REPO" "Resolved — all payroll withdrawals matched"
    rm -f "$INTER" "$FILLED"

    echo ""
fi

# ======================================================================
# REPO 5: Stripe Balance Invariant (single-source)
# ======================================================================
if [[ -z "$ONLY_REPO" || "$ONLY_REPO" == "5" ]]; then
    echo "━━━ REPO 5: nakatomi/stripe-balance-invariant ━━━"
    echo ""

    TEMPLATE="$TEMPLATE_DIR/recon-template.sheet"
    REPO="nakatomi/stripe-balance-invariant"

    # Rev 1: Baseline — balanced ledger
    echo "--- Rev 1: Baseline (PASS) ---"
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$FIXTURE_DIR/ledger-good.csv" \
        --target "tx!A2" --headers --clear \
        --out "$FILLED" --json

    B7=$(read_cell "$FILLED" summary B7)
    echo "  Variance: $B7"

    hub_publish "$FILLED" "$REPO" "Baseline — settled period, balance = 0"
    rm -f "$FILLED"

    # Rev 2: Same data, daily run
    echo "--- Rev 2: Daily run (PASS) ---"
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$FIXTURE_DIR/ledger-good.csv" \
        --target "tx!A2" --headers --clear \
        --out "$FILLED" --json

    hub_publish "$FILLED" "$REPO" "Daily run — balanced, no drift"
    rm -f "$FILLED"

    # Rev 3: Drift — undistributed balance
    echo "--- Rev 3: Drift (FAIL) ---"
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$FIXTURE_DIR/ledger-bad-undistributed.csv" \
        --target "tx!A2" --headers --clear \
        --out "$FILLED" --json

    B7=$(read_cell "$FILLED" summary B7)
    echo "  Variance: $B7 (undistributed)"

    hub_publish "$FILLED" "$REPO" "Drift — \$2,000 undistributed (missing payout for ch_099)"
    rm -f "$FILLED"

    # Rev 4: Resolved
    echo "--- Rev 4: Resolved (PASS) ---"
    FILLED=$(mktemp --suffix=.sheet)

    vgrid fill "$TEMPLATE" \
        --csv "$FIXTURE_DIR/ledger-good.csv" \
        --target "tx!A2" --headers --clear \
        --out "$FILLED" --json

    hub_publish "$FILLED" "$REPO" "Resolved — po_099 settled, balance = 0"
    rm -f "$FILLED"

    echo ""
fi

echo "=== Done ==="
echo ""
echo "View results:"
echo "  https://visihub.app/nakatomi/stripe-qbo-reconciliation"
echo "  https://visihub.app/nakatomi/stripe-mercury-reconciliation"
echo "  https://visihub.app/nakatomi/stripe-brex-reconciliation"
echo "  https://visihub.app/nakatomi/gusto-mercury-payroll"
echo "  https://visihub.app/nakatomi/stripe-balance-invariant"
