# CLI Output Schema — `vgrid publish`

> **Stability: Stable.**
> These JSON keys are a public contract. CI scripts and downstream tools
> parse this output. Breaking changes require a version bump and migration notice.

## Output Modes

| Condition | Output |
|-----------|--------|
| stdout is a TTY | Human-readable text on stderr |
| stdout is piped (CI) | JSON on stdout |
| `--output json` | JSON on stdout (forced) |
| `--output text` | Human-readable text on stderr (forced) |

## JSON Schema: `--wait` (default)

When `--wait` is set (or defaulted), `vgrid publish` waits for the import
to complete and prints the full run result.

```json
{
  "run_id": "42",
  "version": 3,
  "status": "verified",
  "check_status": "pass",
  "diff_summary": {
    "row_count_change": 10,
    "col_count_change": 0
  },
  "row_count": 1000,
  "col_count": 15,
  "content_hash": "blake3:a7ffc6f8...",
  "source_metadata": {
    "type": "dbt",
    "identity": "models/payments",
    "timestamp": "2025-06-15T14:30:00Z",
    "query_hash": "sha256:e3b0c442..."
  },
  "proof_url": "https://api.visihub.app/api/repos/acme/payments/runs/42/proof"
}
```

### Field Reference

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `run_id` | string | **always** | Server-assigned run identifier |
| `version` | integer | **always** | Dataset version number |
| `status` | string | **always** | `"verified"` or `"completed"` |
| `check_status` | string \| null | optional | `"pass"`, `"warn"`, `"fail"`, or `"baseline_created"` (null if checks disabled) |
| `diff_summary` | object \| null | optional | Row/column change summary |
| `row_count` | integer \| null | optional | Total rows in this version |
| `col_count` | integer \| null | optional | Total columns in this version |
| `content_hash` | string \| null | optional | `blake3:<hex>` content hash |
| `source_metadata` | object \| null | optional | Source provenance (see below) |
| `assertions` | array \| null | optional | Control total assertion results (see below) |
| `proof_url` | string | **always** | URL to the cryptographic proof |

### `diff_summary` Fields

| Field | Type | Description |
|-------|------|-------------|
| `row_count_change` | integer | Rows added (positive) or removed (negative) |
| `col_count_change` | integer | Columns added (positive) or removed (negative) |
| `columns_added` | string[] | Names of new columns |
| `columns_removed` | string[] | Names of removed columns |
| `columns_type_changed` | string[] | Names of columns with type changes |

### `source_metadata` Fields

| Field | Type | Description |
|-------|------|-------------|
| `type` | string | Source system (`"dbt"`, `"airflow"`, `"cron"`, etc.) |
| `identity` | string | Model/pipeline identifier |
| `timestamp` | string | ISO 8601 UTC timestamp of the publish |
| `query_hash` | string | Hash of the source query (if applicable) |

### `assertions` Array Items

Each item represents a control total assertion result.

| Field | Type | Description |
|-------|------|-------------|
| `kind` | string | Assertion type (`"sum"`) |
| `column` | string | Column name the assertion applies to |
| `expected` | string \| null | Expected value (strings to avoid float issues) |
| `actual` | string \| null | Computed actual value |
| `tolerance` | string \| null | Allowed deviation |
| `status` | string | `"pass"`, `"fail"`, `"baseline_created"`, or `"skipped"` |
| `delta` | string \| null | Absolute difference (only present on `"fail"`) |
| `message` | string \| null | Error message (only present on errors) |

### CLI Input

```
--assert-sum <column>:<expected>[:<tolerance>]
```

Examples:
- `--assert-sum amount:12345.67:0.01` — sum of `amount` must equal 12345.67 within 0.01
- `--assert-sum revenue:100000` — exact match, zero tolerance

On baseline runs (first revision), if no expected value is given, the actual
value is recorded and status is `"baseline_created"`.

### Check Policy Flags

Control which integrity checks produce `"warn"` vs `"fail"`:

```
--row-count-policy <warn|fail>
--columns-added-policy <warn|fail>
--columns-removed-policy <warn|fail>
--strict                              # sets all policies to fail
```

Defaults: `row_count=warn`, `columns_added=warn`, `columns_removed=fail`.

- `"warn"` means the check detected drift but it's expected. Exit code stays 0.
- `"fail"` means the drift is a real breakage. Exit code 41 (with `--fail-on-check-failure`).
- `--strict` is a macro that sets all three policies to `"fail"`.

Examples:
- `--strict` — break on any drift (frozen extracts)
- `--row-count-policy fail` — break on row count changes but allow column adds
- (no flags) — default: only column removal fails

## JSON Schema: `--no-wait`

When `--no-wait` is set, `vgrid publish` returns immediately after upload
without waiting for server-side processing.

```json
{
  "run_id": "100",
  "status": "processing",
  "proof_url": "https://api.visihub.app/api/repos/acme/payments/runs/100/proof"
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `run_id` | string | **always** | Server-assigned run identifier |
| `status` | string | **always** | Always `"processing"` |
| `proof_url` | string | **always** | URL where the proof will be available after processing |

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success — check passed or warned (or no check configured) |
| 2 | Usage error — bad arguments, missing file |
| 40 | Not authenticated — run `vgrid login` first |
| 41 | Integrity check failed (only with `--fail-on-check-failure`) |
| 42 | Network/HTTP error communicating with VisiHub |
| 43 | Server validation error (bad request) |
| 44 | Timeout waiting for import to complete |

**Note:** Exit code 41 is only returned when `--fail-on-check-failure` is set
AND `check_status` is `"fail"`. A `check_status` of `"warn"` exits 0 (same as
pass). The JSON output is still printed before exiting, so scripts can capture
the `run_id` and `proof_url` even on failure.

## Versioning Policy

- **Additive changes** (new optional fields) are non-breaking and do not require a version bump.
- **Removing a field**, **renaming a field**, or **changing a required field to optional** is a breaking change.
- Breaking changes bump the schema version and are announced in release notes.
- The golden example files in `crates/hub_client/tests/golden/` are the machine-verifiable source of truth.
