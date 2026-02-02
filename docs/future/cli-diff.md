# `visigrid diff` — Reconciliation Spec (v1)

> **Status:** Shipped (v0.3.10).
>
> **Origin:** Real reconciliation session (TJH/Company 7 AP dispute, Jan 2026).
> User had to write a custom Ruby script, deploy it to a server, and run it
> against a database to do what should have been a one-liner.
>
> **Changes to this spec require an RFC.** This is the contract.

---

## Goal

Given two tabular datasets A and B, reconcile them by key:

- rows only in A
- rows only in B
- rows in both but with value differences (with numeric tolerance)
- summary counts and totals

This command is built for AP/AR disputes, invoice matching, and "your list
vs my list" workflows.

**VisiGrid CLI is not trying to be pandas.** It's the tool for reconciliation
and spreadsheet-adjacent operations: convert, compute, compare.

---

## Command

```
visigrid diff <left> <right> [options]
```

`<left>` and `<right>` accept: `.xlsx`, `.csv`, `.tsv`, `.json`, `.sheet`
(same loader contract as `convert`/`calc`).

---

## Required Concepts

### Header row

- Default: first non-empty row is header.
- `--no-headers` treats columns as A, B, C...
- `--header-row N` (1-indexed) overrides detection.

### Data rows

- Data begins the row after header.
- Skip rows that are entirely blank.
- Trim trailing fully-blank columns from bounds (existing bounds logic).

---

## Options

### Key selection

```
--key <col>
```

Column name (if headers) or column letter (A, B, C) or 1-indexed number.

Key values are coerced to string using display representation after transforms.

### Matching mode

```
--match exact|contains
```

**`exact`** (default): Keys must match exactly after transforms.

**`contains`** (v1, caged):
- Left key (after transforms) must be a substring of right key.
  Matching is directional: left → right only. This matches reconciliation
  reality (vendor short ID vs our full ID). Bidirectional contains is a
  different mode if ever needed.
- **Warning:** When `--match contains` is used, a warning is emitted to
  stderr: `warning: using substring matching (--match contains); ensure
  keys are normalized`. This reinforces that exact matching is the default
  and preferred mode.
- Ambiguity policy:
  - 0 right matches → unmatched.
  - 1 right match → matched.
  - \>1 right matches → **ambiguous** → controlled by `--on-ambiguous`.

```
--on-ambiguous error|report
```

- **`error`** (default): exit non-zero, print ambiguous matches to stderr.
- **`report`**: include ambiguous matches in output with status `ambiguous`
  and candidate list. No auto-pairing.

**No "pick best match" in v1. Ever. Not even as an option.**

### Key transforms (v1)

```
--key-transform none|trim|digits
```

Applied in order:
1. `trim` — strip leading/trailing whitespace
2. `digits` — extract only 0–9, drop everything else

Default: `trim`.

Rationale: deterministic, explainable, reduces need for fuzzy matching.

### Duplicate keys

```
--on-duplicate error|group
```

- **`error`** (default): if any duplicate key exists in either dataset,
  exit non-zero and print which keys are duplicated with counts.
- **`group`**: treat duplicate rows under a key as a group (deferred — v1
  ships `error` only).

### Compare columns

```
--compare <col>[,<col>...]
```

- If omitted: compare all columns other than key.
- If provided: compare only these columns.

### Numeric tolerance

```
--tolerance <abs>
```

Default: `0` (exact numeric compare).

Applies only when both sides parse as numbers.

**Numeric parsing rules:**
- Remove commas: `1,234.56` → `1234.56`
- Allow `$` prefix: `$685.00` → `685.00`
- Allow parentheses for negatives: `(123.45)` → `-123.45`
- Allow whitespace
- Treat empty/blank as null

If, after stripping allowed symbols (`$`, `()`, `,`, whitespace), the
remaining value contains non-numeric characters → treat as string.

If one side parses as number and the other doesn't → string mismatch.

### Output format

```
--out json|csv
--output <path>       # default: stdout
--summary none|stderr|json
```

- `--summary none`: suppress stderr summary. JSON output still includes
  `summary` object (it's metadata, always present).
- `--summary stderr` (default): print summary to stderr as human text.
- `--summary json`: include summary object in JSON output (always present
  regardless of this flag; this mode also prints to stderr).

---

## Output Schema

### JSON (`--out json`)

```json
{
  "summary": {
    "left_rows": 109,
    "right_rows": 157,
    "matched": 69,
    "only_left": 40,
    "only_right": 88,
    "diff": 3,
    "ambiguous": 0,
    "tolerance": 0.01,
    "key": "Invoice",
    "match": "contains",
    "key_transform": "trim"
  },
  "results": [
    {
      "status": "only_left",
      "key": "INV-456",
      "left": { "Invoice": "INV-456", "Total": "1200.00" },
      "right": null,
      "diffs": null,
      "match_explain": null,
      "candidates": null
    },
    {
      "status": "diff",
      "key": "16",
      "left": { "Invoice": "16", "Total": "685.00" },
      "right": { "Invoice": "100154662", "Total": "684.99" },
      "diffs": [
        {
          "column": "Total",
          "left": "685.00",
          "right": "684.99",
          "delta": 0.01,
          "within_tolerance": true
        }
      ],
      "match_explain": {
        "mode": "contains",
        "left_key_raw": "16",
        "right_key_raw": "100154662",
        "left_key_norm": "16",
        "right_key_norm": "100154662"
      },
      "candidates": null
    },
    {
      "status": "ambiguous",
      "key": "12",
      "left": { "Invoice": "12", "Total": "500.00" },
      "right": null,
      "diffs": null,
      "match_explain": null,
      "candidates": [
        { "right_key_raw": "100154612", "right_row_index": 42 },
        { "right_key_raw": "100154312", "right_row_index": 87 }
      ]
    }
  ]
}
```

Notes:
- `left` and `right` are null when status is `only_right` / `only_left`.
- `diffs` present only for `diff` status.
- `match_explain` present whenever match mode is not `exact`.
- `candidates` present only for `ambiguous` status.

### CSV (`--out csv`)

Columns:

| Column | Description |
|--------|-------------|
| `status` | `only_left`, `only_right`, `matched`, `diff`, `ambiguous` |
| `key` | Normalized key value |
| `column` | Blank unless `diff` |
| `left_value` | Value from left dataset |
| `right_value` | Value from right dataset |
| `delta` | Numeric only |
| `within_tolerance` | `true`/`false` |
| `match_mode` | `exact`, `contains` |
| `match_explain` | Compact: `contains left="16" right="100154662"` |

Pipeable and grep-friendly.

---

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success (even if diffs exist — diffs are data, not errors) |
| 2 | Usage/args error |
| 3 | Duplicate keys (`--on-duplicate error`) |
| 4 | Ambiguous matches (`--on-ambiguous error`) |
| 5 | Parse/load error |

---

## Concrete Examples

### 1. Exact match by invoice

```bash
visigrid diff vendor.xlsx ours.csv \
  --key Invoice \
  --compare Total \
  --tolerance 0.01 \
  --out json
```

### 2. Contains match for short IDs inside long IDs

```bash
visigrid diff her.xlsx our_export.csv \
  --key Order \
  --match contains \
  --on-ambiguous error \
  --compare Amount \
  --tolerance 0.01 \
  --out csv
```

### 3. Digits-only key normalization

```bash
visigrid diff her.xlsx our_export.csv \
  --key Order \
  --key-transform digits \
  --compare Amount \
  --tolerance 0.01
```

### Summary output (stderr)

```
left:  109 rows (vendor.xlsx)
right: 157 rows (ours.csv)
matched: 69
only_left: 40
only_right: 88
value_diff: 3 (Total, net delta: $12.47)
```

---

## Test Matrix

All 10 golden tests passing. Test data in `tests/cli/diff/`.

| Test | Asserts | Status |
|------|---------|--------|
| `exact_match_no_diffs` | matched count, exit 0 | PASS |
| `exact_only_left` | key in A missing in B | PASS |
| `exact_only_right` | key in B missing in A | PASS |
| `exact_diff_tolerance_pass` | delta within tolerance → `within_tolerance: true` | PASS |
| `exact_diff_tolerance_fail` | delta exceeds tolerance → `within_tolerance: false` | PASS |
| `contains_single_match` | 1:1 substring match succeeds with explain | PASS |
| `contains_ambiguous_errors` | >1 match → exit 4, prints candidates | PASS |
| `duplicate_keys_error_left` | duplicate in A → exit 3 with details | PASS |
| `currency_parsing` | `$1,234.56` and `(500.00)` parse correctly | PASS |
| `blank_row_skip` | blank rows excluded from data bounds | PASS |

---

## Architecture

### Module boundary

```
crates/cli/src/diff.rs    # DiffOptions + DiffResult + diff logic
crates/cli/src/main.rs    # Clap variant + IO + formatting
```

`diff.rs` takes two `Vec<DataRow>` + headers + `DiffOptions` → `DiffResult`.
No IO, no formatting. Pure reconciliation logic. `DataRow` contains the
raw key, normalized key, and column values as `HashMap<String, String>`.

CLI layer handles clap parsing, calling `read_file()`, formatting output,
writing to stdout/stderr.

### Reused code paths

- `read_file()` / `read_stdin()` — format detection, all loaders
- `get_data_bounds()` — find data extent
- `write_csv()` / `write_json()` patterns — output formatting
- `CliError` — exit code routing

### No new dependencies

Engine and IO crates unchanged. No new crate deps in CLI.

---

## Deferred (post-v1)

| Feature | Reason |
|---------|--------|
| `--on-duplicate group` | Group-level comparison needs aggregation rules |
| `--match suffix` | Covered by `--key-transform digits` + exact |
| `--match regex` | Danger zone — hard to explain, easy to misuse |
| `--tolerance N%` | Percentage tolerance adds edge cases |
| `--key-transform` chaining | Multiple transforms in sequence |
| `--key-left` / `--key-right` | Column name mapping between files |
| Composite keys | `--key "Last,First"` |
| `--where` filtering | Scope creep toward query engine |
| `--columns` selection | Useful but independent of diff |
