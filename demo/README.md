[![Nightly Check](https://github.com/VisiGrid/demo/actions/workflows/nightly.yml/badge.svg)](https://github.com/VisiGrid/demo/actions/workflows/nightly.yml)

# VisiGrid

**If your finance data changes silently, this catches it.**

Treat finance data like code: versioned, reviewed, enforced, and cryptographically proven.

## One Command

```bash
curl -fsSL https://get.visigrid.app/install.sh | sh

./demo/scripts/stripe_recon.sh \
  --repo YOUR_ORG/YOUR_REPO \
  --from 2026-01-01 --to 2026-02-01
```

Fetches your Stripe balance transactions, fills a recon template, computes the undistributed balance, publishes a signed proof. If the invariant breaks, CI fails.

Three modes:

| Mode | What it does |
|------|-------------|
| `--mode baseline` | Fetch + publish. No assertion. Use for initial runs or unsettled windows. **(default)** |
| `--mode assert-zero` | Fetch + publish + assert B7 = 0. Use for fully settled windows. |
| `--mode break` | Skip fetch, inject bad fixture data. Proves the assertion catches real violations. |

## What B7 Means

Cell `summary!B7` is the **undistributed balance** — the sum of all Stripe balance transaction amounts across every category:

```
B7 = Charges + Payouts + Fees + Refunds + Adjustments
```

Stripe encodes direction in sign: charges are positive, payouts/fees/refunds are negative. For a fully settled period, everything that came in has gone out and B7 = 0.

| B7 value | Meaning |
|----------|---------|
| `0` | Balanced. Every dollar in has a dollar out. |
| Positive (e.g. `17500`) | $175.00 processed but not yet paid out. Normal for recent windows — payouts lag 2-7 days. |
| Negative | More paid out than charged. Unusual — investigate refunds or adjustments. |

**Why it works as a CI gate:** For a settled window (e.g. last month after all payouts cleared), B7 must be 0. If it's not, something changed — a retroactive refund, a missing fee, a payout that didn't match. The assertion catches it, signs a proof, and fails the pipeline.

## Raw Commands

The script runs three `vgrid` commands. Here they are for direct use:

```bash
# 1. Fetch Stripe balance transactions → canonical CSV
vgrid fetch stripe \
  --from 2026-01-01 --to 2026-02-01 \
  --out stripe.csv

# 2. Fill the recon template with fetched data
vgrid fill demo/templates/recon-template.sheet \
  --csv stripe.csv \
  --target "tx!A2" --headers --clear \
  --out recon.sheet

# 3. Publish with invariant assertion
vgrid publish recon.sheet \
  --repo YOUR_ORG/YOUR_REPO \
  --dataset ledger-recon \
  --source-type stripe \
  --wait \
  --assert-cell "summary!B7:0:0" \
  --output json
```

Read B7 without publishing:

```bash
vgrid sheet inspect recon.sheet --sheet summary B7 --json
```

## GitHub Actions

```yaml
- name: Install vgrid
  run: |
    curl -fsSL https://get.visigrid.app/install.sh | sh
    echo "$HOME/.local/bin" >> $GITHUB_PATH

- name: Stripe recon
  run: |
    vgrid login --token "${{ secrets.VISIHUB_API_KEY }}"
    ./demo/scripts/stripe_recon.sh \
      --repo "${{ vars.VISIHUB_REPO }}" \
      --from "$(date -d '45 days ago' +%Y-%m-%d)" \
      --to "$(date +%Y-%m-%d)" \
      --mode assert-zero \
      --api-key "${{ secrets.STRIPE_API_KEY }}"
```

Job fails if undistributed balance != 0. Proof is signed and stored regardless.

## Check Policy

| Change | Default | Effect |
|--------|---------|--------|
| Row count | warn | Exit 0 |
| Columns added | warn | Exit 0 |
| Columns removed | **fail** | Exit 41 |
| Invariant violated | **fail** | Exit 41 |

Override: `--row-count-policy fail` or `--strict`.

## Fixtures

| File | Scenario |
|------|----------|
| `data/ledger-good.csv` | 4 balanced rows (net zero) |
| `data/ledger-bad-append.csv` | +1 row |
| `data/ledger-bad-remove.csv` | -1 row |
| `data/ledger-bad-column-add.csv` | +1 column |
| `data/ledger-bad-undistributed.csv` | $2,000 unmatched charge (B7 = 200000) |
| `templates/recon-template.sheet` | SUMIF recon with variance formula at B7 |

## What This Is

- Dataset version control
- Structural drift detection
- Financial invariant enforcement
- Signed proof registry
- CI-native

## What This Is Not

- ETL
- BI
- A reconciliation UI
- A connector marketplace

You bring the data. VisiGrid verifies it hasn't drifted and signs the proof.

## Pricing

Automated publishing + proof retention require **Professional ($79/mo)**.
Free tier is for manual verification.

[visihub.app](https://visihub.app)
