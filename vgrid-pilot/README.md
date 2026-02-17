# vgrid recon — 1-Week Pilot Kit

`vgrid recon` is a deterministic reconciliation engine that compares your processor (Stripe), ledger (NetSuite/QBO), and bank (Mercury/SVB/etc.) exports and tells you exactly what matches, what's off, and by how much — before month-end.

## Quick Start (2 minutes)

### 1. Verify install

```
vgrid --version
# Expected: vgrid 0.8.3 or later
```

If you don't have it yet: `curl -fsSL https://get.visigrid.app | sh`

### 2. Run the demo (clean match)

```
vgrid recon run demo/clean-match.recon.toml
```

Expected output:
```
2-way recon: 4 groups — 4 matched, 0 amount mismatches, 0 timing mismatches, 0 unmatched
```

### 3. Run the demo (deliberate mismatch)

```
vgrid recon run demo/has-issues.recon.toml
```

Expected output:
```
2-way recon: 4 groups — 2 matched, 0 amount mismatches, 0 timing mismatches, 2 unmatched
```

### 4. See JSON detail

```
vgrid recon run demo/has-issues.recon.toml --json
```

Every group gets a bucket: `matched_two_way`, `amount_mismatch`, `timing_mismatch`, `ledger_only`, `processor_ledger_only`, etc.

---

## What You Need to Provide (3 files)

| File | What it is | Where to get it |
|------|-----------|----------------|
| **processor.csv** | Stripe payouts or balance transactions | Stripe Dashboard → Payouts → Export |
| **ledger.csv** | Deposits / cash receipts from your ERP | NetSuite Saved Search, QBO Reports, Xero export |
| **bank.csv** (optional, for 3-way) | Bank deposits | Mercury/SVB/Chase CSV download |

**Requirements:**
- CSV format, UTF-8
- Must have columns for: unique ID, date, amount (integer cents or dollars), currency, transaction type
- Column names don't matter — we map them in the config

**Privacy:** You can redact descriptions and counterparty names. The engine only needs: ID, amount, date, currency, type. You can also multiply all amounts by a constant if you want to obscure real numbers — relative matching still works.

---

## Setting Up Your Config (5 minutes)

Copy `recon.example.toml` and edit it for your columns:

```
cp recon.example.toml my-recon.toml
```

The key decisions:

1. **`match_key`** — What ties a processor row to a ledger row?
   - If your Stripe payout ID appears in both files → use it as `match_key` on both sides (exact_key strategy)
   - If it doesn't → use `fuzzy_amount_date` strategy (matches by amount + date window)

2. **`filter`** — Which rows matter?
   - Stripe: usually filter to `type = "payout"` only
   - Ledger: usually filter to `type = "deposit"` or `category = "Stripe"`

3. **`transform`** — Sign flips
   - Stripe payouts are negative (money leaving Stripe). Set `multiply = -1` so they become positive for comparison.

Then run:

```
vgrid recon run my-recon.toml --json --output result.json
```

---

## Exit Codes (for CI/scripts)

| Code | Meaning |
|------|---------|
| 0 | All groups matched within tolerance |
| 1 | Mismatches found (amount, timing, or unmatched) |
| 2 | Runtime error (missing file, bad CSV) |
| 60 | Invalid config |

---

## What the Output Looks Like

```json
{
  "meta": {
    "config_name": "Stripe ↔ QBO Payout Recon",
    "way": 2,
    "engine_version": "0.8.3",
    "run_at": "2026-02-17T15:11:49Z"
  },
  "summary": {
    "total_groups": 47,
    "matched": 45,
    "amount_mismatches": 1,
    "timing_mismatches": 1,
    "left_only": 0,
    "right_only": 0
  },
  "groups": [
    {
      "bucket": "amount_mismatch",
      "match_key": "po_3847",
      "currency": "USD",
      "aggregates": {
        "processor": { "total_cents": 154723, "record_count": 1 },
        "ledger": { "total_cents": 154700, "record_count": 1 }
      },
      "deltas": { "delta_cents": 23, "date_offset_days": 0 }
    }
  ]
}
```

---

## Sending Results Back

Email or Slack the `result.json` file. If you want to redact:

- Amounts are in integer cents — multiply by a constant in your CSV before running, or just send the summary section
- Match keys are your internal IDs — redact if needed
- The `summary` section alone is useful: it tells us matched vs. mismatched counts

---

## Questions?

Reply to this thread or email [your email]. I can build your `recon.toml` config in 10 minutes if you send me a sample CSV header row (no data needed, just the column names).
