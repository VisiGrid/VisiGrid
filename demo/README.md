[![Nightly Check](https://github.com/VisiGrid/demo/actions/workflows/nightly.yml/badge.svg)](https://github.com/VisiGrid/demo/actions/workflows/nightly.yml)

# VisiGrid

**If your finance data changes silently, this catches it.**

Treat finance data like code: versioned, reviewed, enforced, and cryptographically proven.

## 60-Second Integrity Test

```bash
# Install
curl -fsSL https://get.visigrid.app/install.sh | sh

# Publish baseline (4 balanced rows — charges, fees, payout net to zero)
vgrid publish data/ledger-good.csv \
  --repo YOUR_ORG/YOUR_REPO \
  --dataset ledger-recon \
  --source-type manual \
  --wait --output json
# → check_status: "baseline_created"

# Publish with an extra row appended
vgrid publish data/ledger-bad-append.csv \
  --repo YOUR_ORG/YOUR_REPO \
  --dataset ledger-recon \
  --source-type manual \
  --wait --output json
# → check_status: "warn", row_count_change: 1, exit 0
```

Row added? Warning. Row removed? Hard fail. Column disappeared? Hard fail.
No config required.

## The Invariant That Protects Money

```bash
vgrid publish recon.sheet \
  --assert-cell "summary!B7:0:0" \
  --wait
```

If undistributed balance != 0:

- `check_status: "fail"`
- `exit 41`
- Proof signed and stored
- CI fails

```json
{
  "assertions": [{
    "column": "summary!B7",
    "expected": "0",
    "actual": "200000",
    "status": "fail",
    "delta": "200000.0",
    "origin": "client"
  }]
}
```

[Live proof example](https://api.visihub.app/api/repos/robert/test-repo/runs/50/proof)

## GitHub Actions

```yaml
- name: Install vgrid
  run: |
    curl -fsSL https://get.visigrid.app/install.sh | sh
    echo "$HOME/.local/bin" >> $GITHUB_PATH

- name: Publish and check
  run: |
    vgrid login --token "${{ secrets.VISIHUB_API_KEY }}"
    vgrid publish data/export.csv \
      --repo "${{ vars.VISIHUB_REPO }}" \
      --dataset nightly-recon \
      --source-type manual \
      --wait
```

Job fails if data drifts beyond policy. See [`.github/workflows/`](.github/workflows/) for complete examples.

## Check Policy

| Change | Default | Effect |
|--------|---------|--------|
| Row count | warn | Exit 0 |
| Columns added | warn | Exit 0 |
| Columns removed | **fail** | Exit 41 |
| Invariant violated | **fail** | Exit 41 |

Override: `--row-count-policy fail` or `--strict`.

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

## Fixtures

| File | Scenario |
|------|----------|
| `data/ledger-good.csv` | 4 balanced rows (net zero) |
| `data/ledger-bad-append.csv` | +1 row |
| `data/ledger-bad-remove.csv` | -1 row |
| `data/ledger-bad-column-add.csv` | +1 column |
| `data/ledger-bad-undistributed.csv` | $2,000 unmatched charge |
| `templates/recon-template.sheet` | SUMIF recon with variance formula |

See [DEMO.md](DEMO.md) for full receipts with commands, outputs, and proof URLs.

## Pricing

Automated publishing + proof retention require **Professional ($79/mo)**.
Free tier is for manual verification.

[visihub.app](https://visihub.app)
