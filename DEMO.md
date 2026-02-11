# VisiGrid + VisiHub: Abuse Test Receipts

Tested: 2026-02-11
Repo: `robert/test-repo`
Dataset: `abuse-recon-1770796419`
Engine: visigrid-engine 0.6.6

## Publish Tests (#1-#5)

All tests use `crates/cli/tests/abuse/csv/` fixtures against the same
dataset to show the full baseline-pass-drift lifecycle.

### Test #1: Baseline Creation

New dataset, 4 balanced transaction rows (charges + fees + payout = 0).

```
vgrid publish crates/cli/tests/abuse/csv/balanced.csv \
  --repo robert/test-repo \
  --dataset abuse-recon-1770796419 \
  --source-type manual --wait --output json
```

```json
{
  "run_id": "37",
  "version": 2,
  "status": "verified",
  "check_status": "pass",
  "diff_summary": {"col_count_change": 0, "row_count_change": 0},
  "row_count": 4,
  "col_count": 10,
  "content_hash": "blake3:cd8a70f1a001df2a4247de46f2a0e9aa0b51be13134e3b38a25e196528f29689",
  "proof_url": "https://api.visihub.app/api/repos/robert/test-repo/runs/37/proof"
}
```

Exit code: **0**
Baseline version (run 36) auto-created on first upload; run 37 is
identical content → `check_status: "pass"`.

---

### Test #2: Repeat Identical

Same CSV, same dataset. Content hash matches — no drift.

```
vgrid publish crates/cli/tests/abuse/csv/balanced.csv \
  --repo robert/test-repo \
  --dataset abuse-recon-1770796419 \
  --source-type manual --wait --output json
```

```json
{
  "run_id": "38",
  "version": 3,
  "status": "verified",
  "check_status": "pass",
  "diff_summary": {"col_count_change": 0, "row_count_change": 0},
  "row_count": 4,
  "col_count": 10,
  "content_hash": "blake3:cd8a70f1a001df2a4247de46f2a0e9aa0b51be13134e3b38a25e196528f29689",
  "proof_url": "https://api.visihub.app/api/repos/robert/test-repo/runs/38/proof"
}
```

Exit code: **0**
Content hash identical. No structural changes.

---

### Test #3: Append Rows (+1 row)

`balanced-plus-row.csv` has 5 rows (original 4 + new charge).
Row count increased by 1.

```
vgrid publish crates/cli/tests/abuse/csv/balanced-plus-row.csv \
  --repo robert/test-repo \
  --dataset abuse-recon-1770796419 \
  --source-type manual --wait --row-count-policy warn --output json
```

```json
{
  "run_id": "39",
  "version": 4,
  "status": "verified",
  "check_status": "fail",
  "diff_summary": {"col_count_change": 0, "row_count_change": 1},
  "row_count": 5,
  "col_count": 10,
  "content_hash": "blake3:ee8dc1407f1cd75ef11d7f0a1f9160f383ee0c40ce9e2338ff3e2341f4494b4c",
  "proof_url": "https://api.visihub.app/api/repos/robert/test-repo/runs/39/proof"
}
```

Exit code: **41** (EXIT_HUB_CHECK_FAILED)

**Note:** `--row-count-policy warn` was passed but the server returned
`check_status: "fail"`. The check_policy system (default: `row_count:
"warn"`) exists in the codebase but hasn't been deployed to production
yet. After deployment, this will return `check_status: "warn"` and exit
code 0.

---

### Test #4: Remove Rows (-2 rows)

`balanced-minus-row.csv` has 3 rows (removed payout row).
Row count decreased by 2 (vs the 5-row version 4).

```
vgrid publish crates/cli/tests/abuse/csv/balanced-minus-row.csv \
  --repo robert/test-repo \
  --dataset abuse-recon-1770796419 \
  --source-type manual --wait --no-fail --output json
```

```json
{
  "run_id": "40",
  "version": 5,
  "status": "verified",
  "check_status": "fail",
  "diff_summary": {"col_count_change": 0, "row_count_change": -2},
  "row_count": 3,
  "col_count": 10,
  "content_hash": "blake3:034313e1bdcb1825e5e3d843153be33205705cf6c0fd7b4818b059f974b5fcc5",
  "proof_url": "https://api.visihub.app/api/repos/robert/test-repo/runs/40/proof"
}
```

Exit code: **0** (--no-fail suppresses exit code)

Row count dropped by 2. Server correctly flags `check_status: "fail"`.

---

### Test #5: Add Column (+1 column)

`extra-column.csv` has 11 columns (original 10 + `notes`).
Same 4 rows but new column added.

```
vgrid publish crates/cli/tests/abuse/csv/extra-column.csv \
  --repo robert/test-repo \
  --dataset abuse-recon-1770796419 \
  --source-type manual --wait --no-fail --output json
```

```json
{
  "run_id": "41",
  "version": 6,
  "status": "verified",
  "check_status": "fail",
  "diff_summary": {"col_count_change": 1, "row_count_change": 1},
  "row_count": 4,
  "col_count": 11,
  "content_hash": "blake3:718476448ba047fe24315a36eb1f0416fa24e82934117877e3bf5672917c1bc5",
  "proof_url": "https://api.visihub.app/api/repos/robert/test-repo/runs/41/proof"
}
```

Exit code: **0** (--no-fail)

Column count increased by 1, row_count_change 1 (vs the 3-row v5).
Server flags `check_status: "fail"`. After check_policy deployment,
columns_added will report "warn" per the default policy.

---

## Local Abuse Tests (#6-#14)

All 11 local tests pass after three engine/CLI bug fixes.

```
cargo test -p visigrid-cli --test abuse_tests -- --nocapture
```

```
test result: ok. 11 passed; 0 failed; 0 ignored
```

| # | Test | Result |
|---|------|--------|
| 6 | Currency symbol rejection (`$1000.00`) | PASS — exit 4, "currency symbol" |
| 7 | Comma rejection (`1,000`) | PASS — exit 4, "comma in numeric field" |
| 8 | Wrong decimals (`1000.5`) | PASS — exit 4, "wrong decimal places" |
| 9 | Formula injection (`=1+1`) | PASS — exit 4, "formula injection rejected" |
| 10 | Missing column (no amount_minor) | PASS — fill succeeds, SUMIF returns 0 |
| 11a | Stale rows without --clear | PASS — variance != 0 (corruption detected) |
| 11b | Clean fill with --clear | PASS — variance = 200000 (only charge data) |
| 12 | Blank assertion cell | PASS — B10 = Empty |
| 13 | Error assertion cell (=1/0) | PASS — B8 = Error("#DIV/0!") |
| 14 | String assertion cell | PASS — B9 = Text("not a number") |
| -- | Balanced fill variance zero | PASS — charges=150000, variance=0 |

### Test #11 (Stale Rows) — Detailed

The stale-rows scenario is the most important abuse case for financial
reconciliation. Without `--clear`, a shorter CSV leaves old rows behind,
corrupting SUMIF totals silently.

**11a: Without --clear (corruption)**
1. Fill template with balanced.csv (4 rows, variance=0)
2. Fill same file with short.csv (1 row, charge 200000) WITHOUT --clear
3. Result: variance = 100000 (stale fee/payout rows from step 1 remain)

**11b: With --clear (correct)**
1. Fill template with balanced.csv (4 rows)
2. Fill same file with short.csv (1 row) WITH --clear
3. Result: variance = 200000 (only the single charge, no stale data)

---

## Bugs Found and Fixed

### Bug 1: `vgrid fill --headers` wrote CSV headers as data
**File:** `crates/cli/src/fill.rs`
The `--headers` flag parsed the header row but still wrote it to the
sheet. Headers now correctly skipped.

### Bug 2: Leaf formulas not evaluated after save/load
**Files:** `crates/engine/src/dep_graph.rs`, `crates/engine/src/workbook.rs`
Formulas with no cell references (e.g., `=1/0`, `=PI()`) were excluded
from the dependency graph, so `recompute_full_ordered()` never
re-evaluated them after clearing the cache. Added
`register_leaf_formula()` to track these in topo sort.

### Bug 3: Cross-sheet SUMIF/COUNTIF/AVERAGEIF ignored sheet reference
**File:** `crates/engine/src/formula/eval_conditional.rs`
`extract_range()` discarded the sheet field from range expressions.
All conditional functions read from the current sheet instead of the
referenced sheet. Complete rewrite to propagate `SheetRef` through
`range_get_text()` / `range_get_value()` helpers. All 7 conditional
functions fixed (SUMIF, AVERAGEIF, COUNTIF, COUNTBLANK, SUMIFS,
AVERAGEIFS, COUNTIFS).

---

## Check Policy Contract

Current default policy (in codebase, pending deployment):

| Check | Default Severity | Meaning |
|-------|-----------------|---------|
| `row_count` | `warn` | Row count changed from baseline |
| `columns_added` | `warn` | New columns appeared |
| `columns_removed` | `fail` | Columns disappeared |

**Current production behavior:** All drift = `fail` (check_policy
migration not yet deployed). After deployment, row_count and
columns_added drift will produce `warn` status and exit code 0 (unless
`--strict` is used).

**Planned enhancement:** Separate `row_count_increase` and
`row_count_decrease` policies. Row count increase (append) is lower
severity than decrease (data loss). Not yet implemented.

### Exit code behavior

| check_status | --fail-on-check-failure (default) | --no-fail |
|-------------|----------------------------------|-----------|
| pass | exit 0 | exit 0 |
| warn | exit 0 | exit 0 |
| fail | exit 41 | exit 0 |
| baseline_created | exit 0 | exit 0 |

---

## Cell Assertion (--assert-cell)

The `--assert-cell` feature is implemented in the CLI and hub_client.
Server-side opaque file handling (`finalize!` for .sheet files) is
implemented in code but not deployed. Once deployed:

```
vgrid publish recon.sheet \
  --repo acme/payments \
  --assert-cell "summary!B7:0:10000" \
  --wait
```

This evaluates `summary!B7` locally using the VisiGrid engine, sends
the actual value with `origin: "client"` and engine metadata
(name, version, fingerprint), and the server performs the
comparison and signs the proof.
