# `visigrid diff` — Behavioural Contract

**contract_version: 1**

This document is the stable reference for `visigrid diff` JSON output semantics.
The original design spec lives in `docs/future/cli-diff.md`; this file records
what the code actually guarantees.

---

## JSON schema

Every JSON invocation (`--out json`) produces a top-level object with exactly
three keys in this order:

```json
{
  "contract_version": 1,
  "results": [ ... ],
  "summary": { ... }
}
```

- **`contract_version`** — integer. Incremented when a backwards-incompatible
  change is made to this schema. Consumers should assert on this field.
- **`results`** — array of row-result objects (see Row Statuses below).
- **`summary`** — object with aggregate counts and option echo.

Key ordering within each object is alphabetical (serde_json BTreeMap).
This is an implementation detail, not a guarantee — consumers must not depend
on key order.

### Summary fields

| Field | Type | Meaning |
|---|---|---|
| `left_rows` | int | Total data rows in left file |
| `right_rows` | int | Total data rows in right file |
| `matched` | int | Rows matched with identical compared values |
| `diff` | int | Rows matched but with value differences |
| `diff_outside_tolerance` | int | Subset of `diff` where at least one delta exceeds tolerance |
| `only_left` | int | Left rows with no right match |
| `only_right` | int | Right rows with no left match |
| `ambiguous` | int | Left rows with multiple right candidates (only with `--on-ambiguous report`) |
| `tolerance` | float | Echo of `--tolerance` option |
| `key` | string | Echo of `--key` column name |
| `match` | string | Echo of match mode: `"exact"` or `"contains"` |
| `key_transform` | string | Echo of key transform: `"none"`, `"trim"`, or `"digits"` |

---

## Row statuses

Each element of `results` has a `status` field:

| Status | Meaning |
|---|---|
| `matched` | Key found on both sides, all compared values equal |
| `diff` | Key found on both sides, at least one compared value differs |
| `only_left` | Key present in left file, no match in right file |
| `only_right` | Key present in right file, no match in left file |
| `ambiguous` | Key present in left file, multiple right candidates (reported when `--on-ambiguous report`) |

---

## Tolerance semantics

Tolerance applies only when **both** values parse as financial numbers
(see Numeric Parsing below). Non-numeric values are compared as byte-exact
strings; tolerance has no effect on them.

- **Inclusive boundary**: a delta is within tolerance when
  `delta <= tolerance + epsilon`, where epsilon is a scale-aware IEEE-754
  adjustment: `eps = f64::EPSILON * 16.0 * scale`,
  `scale = max(1.0, |left|, |right|, |delta|, |tolerance|)`.
- **Why epsilon**: `0.01` tolerance for `100.50 vs 100.49` must pass.
  Naïve `f64` subtraction produces `0.01000000000000512…`, which fails
  a strict `<=` check. The epsilon band corrects for representation error
  at the magnitude in play.
- **`within_tolerance`**: each column diff reports this boolean. `true` means
  the delta is within the tolerance band (inclusive of epsilon). A row with
  only within-tolerance diffs has status `diff` and increments `summary.diff`
  but **not** `summary.diff_outside_tolerance`.
- **Zero tolerance**: when `--tolerance 0` (default), epsilon is still applied.
  Two numerically identical values parsed from different string representations
  (e.g. `$1,000.00` vs `1000`) produce `delta = 0.0` and `within_tolerance = true`.

---

## Exit codes

| Code | Meaning |
|---|---|
| **0** | Fully reconciled. All left rows matched, no only-right rows, no diffs outside tolerance. |
| **1** | Material drift. Missing rows (`only_left > 0` or `only_right > 0`) or value diffs outside tolerance. With `--strict-exit`, any diff (even within tolerance) triggers exit 1. |
| **≥ 2** | Error. Argument validation, I/O, parse, duplicate keys, ambiguous keys, etc. Specific codes in this range are best-effort and may change between releases. |

### `--strict-exit`

When enabled, **any** row with status `diff` (regardless of tolerance) causes
exit 1. Without this flag, within-tolerance diffs are reported but do not
cause a non-zero exit code.

---

## Determinism

Same input files + same options → byte-identical JSON output.

Row ordering:

1. All left rows in their original file order (matched, diff, ambiguous, only_left)
2. Remaining only-right rows in their original right-file order

---

## Numeric parsing

`parse_financial_number` strips the following before parsing as `f64`:

- Leading/trailing whitespace
- `$` currency prefix
- `,` thousands separators
- Parenthesized negatives: `(123.45)` → `-123.45`

After stripping, only `0-9`, `.`, and a leading `-`/`+` may remain.
Anything else → value is treated as a non-numeric string.

Examples:

| Input | Parsed as |
|---|---|
| `$1,234.56` | `1234.56` |
| `(500.00)` | `-500.0` |
| ` 42 ` | `42.0` |
| `N/A` | non-numeric (string compare) |
| `` (empty) | skip (both-empty = match) |

---

## Unicode

All string comparisons are **byte-exact**. No NFC/NFD normalization is
performed. If the left file encodes `é` as U+00E9 and the right file uses
U+0065 U+0301, they compare as different strings.

---

## Key transforms

| Transform | Effect |
|---|---|
| `none` | Key used as-is (after trimming leading/trailing whitespace) |
| `trim` | Same as `none` (trim is the default) |
| `digits` | Extract only ASCII digit characters from the key |

Transforms are applied to both left and right keys before matching.

---

## CSV output

CSV output (`--out csv`) does **not** include `contract_version` or the
summary object. It produces one row per result with columns:
`status`, `key`, `column`, `left`, `right`, `delta`, `within_tolerance`.
CSV format is not covered by this contract version and may evolve.

---

## Backward compatibility

- **Within a contract version**: JSON output shape, field names, status
  strings, tolerance semantics, and exit code classes (0/1/≥2) are stable.
  New fields may be **added** to objects; existing fields will not be removed
  or change type.
- **What triggers a version bump**: removing a field, renaming a field,
  changing the type of a field, changing tolerance arithmetic, or changing
  exit code class semantics (e.g. making exit 0 mean something other than
  "fully reconciled").
- **Deprecation**: best-effort. When a breaking change is planned, the prior
  version's behavior will be documented in a release note. There is no
  formal multi-version support; `contract_version` tells you which contract
  is in effect.
- **Specific exit codes ≥ 2**: these are best-effort diagnostics and may be
  reassigned between releases. Do not match on specific codes above 1;
  treat `≥ 2` as "error."
