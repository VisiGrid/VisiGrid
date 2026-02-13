# Lua Scripts & Run Records — Schema Specification

## Design Principles

1. **Scripts are metadata, not data.** Attached scripts are excluded from the semantic fingerprint. The fingerprint covers cell values, formulas, and cell metadata only.
2. **Never auto-execute.** Opening a .sheet with attached scripts does nothing. Execution requires explicit user action.
3. **Content-addressed.** Every script is identified by `sha256(source)`. This enables Hub verification, deduplication, and tamper detection.
4. **Run records snapshot everything.** Each execution stores the full script source used at that time. Reproducibility doesn't depend on file history.
5. **Schema versioned from day one.** All JSON structures include `schema_version: 1`.

---

## 1. Script Metadata

The canonical shape of a script, whether attached, project, or global.

```json
{
  "schema_version": 1,
  "name": "reconcile_vendors",
  "description": "Normalize vendor names against master list",
  "hash": "sha256:e3b0c44298fc1c149afb...",
  "source": "for r = 2, sheet:rows() do\n  ...\nend",
  "capabilities": ["sheet.read", "sheet.write.values"],
  "created_at": "2026-02-12T14:30:00Z",
  "updated_at": "2026-02-12T14:30:00Z",
  "origin": null,
  "author": null,
  "version": null
}
```

### Fields

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `schema_version` | integer | yes | Always `1` |
| `name` | string | yes | Human-readable name (also used as identifier in resolution) |
| `description` | string | no | One-line description shown in palette |
| `hash` | string | yes | `sha256:<hex>` of the `source` field. Recomputed on save. |
| `source` | string | yes | Full Lua source code |
| `capabilities` | string[] | yes | Declared capabilities (see below) |
| `created_at` | string | yes | ISO 8601 timestamp |
| `updated_at` | string | yes | ISO 8601 timestamp |
| `origin` | string | no | Where this script came from. `null` = local. Future: `"hub:org/name@v1"` |
| `author` | string | no | Reserved for Hub. Human-readable author identifier. |
| `version` | string | no | Reserved for Hub. Semver or revision number. |

### Capabilities

Capabilities use a namespaced enum that can grow without breaking existing scripts.

**Phase 1 (now):**

| Capability | Meaning | Detected by |
|------------|---------|-------------|
| `sheet.read` | Reads cell values/formulas: `get_value`, `get_formula`, `range:values()`, `selection()`, `rows()`, `cols()` | Presence of read API calls |
| `sheet.write.values` | Writes cell values: `set_value`, `range:set_values()`, `clear` | Presence of write API calls |
| `sheet.write.formulas` | Writes formulas: `set_formula` | Presence of `set_formula` calls |

**Reserved (Phase 2+):**

| Capability | Meaning |
|------------|---------|
| `sheet.write.format` | Modifies cell formatting (if ever exposed to Lua) |
| `clipboard.read` | Reads system clipboard |
| `clipboard.write` | Writes to system clipboard |
| `fs.read` | Reads external files |
| `fs.write` | Writes external files |
| `net.http` | Makes HTTP requests |

A script with `["sheet.read"]` only inspects data. A script with `["sheet.read", "sheet.write.values"]` modifies cell values. A script with `[]` is a pure computation (math, string ops, print output).

Capabilities are declared at save time by static analysis of the source (presence of API calls matching each capability). The sandbox enforces declared capabilities at runtime — a script declaring only `sheet.read` that attempts `set_value` will error.

The namespace convention (`domain.verb.target`) allows fine-grained capability growth without breaking the schema or existing scripts.

---

## 2. Attached Scripts (.sheet storage)

.sheet files are SQLite databases. Attached scripts are stored in a new `scripts` table.

### SQL schema (migration to schema version 8)

```sql
CREATE TABLE IF NOT EXISTS scripts (
    name TEXT PRIMARY KEY,
    hash TEXT NOT NULL,
    source TEXT NOT NULL,
    description TEXT,
    capabilities TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    origin TEXT,
    author TEXT,
    version TEXT
);
```

`capabilities` is stored as a JSON array string (e.g., `'["sheet.read","sheet.write.values"]'`).

### Fingerprint exclusion

The `scripts` table is explicitly excluded from fingerprint computation. The fingerprint function reads from `cells` and `cell_metadata` only. This is already the case — new tables are ignored by the fingerprint unless explicitly included.

### UI indicator

When a .sheet has rows in the `scripts` table, the status bar shows:

```
3 scripts attached
```

Not a warning. Not a trust dialog. Just visibility.

---

## 3. Run Records (.sheet storage)

Each script execution produces a run record stored in the .sheet file.

### SQL schema (same migration, schema version 8)

```sql
CREATE TABLE IF NOT EXISTS run_records (
    run_id TEXT PRIMARY KEY,
    script_name TEXT NOT NULL,
    script_hash TEXT NOT NULL,
    script_source TEXT NOT NULL,
    script_origin TEXT NOT NULL DEFAULT '{}',
    capabilities_used TEXT NOT NULL DEFAULT '[]',
    params TEXT,
    fingerprint_before TEXT NOT NULL,
    fingerprint_after TEXT NOT NULL,
    diff_hash TEXT,
    diff_summary TEXT,
    cells_read INTEGER NOT NULL DEFAULT 0,
    cells_modified INTEGER NOT NULL DEFAULT 0,
    ops_count INTEGER NOT NULL DEFAULT 0,
    duration_ms INTEGER NOT NULL DEFAULT 0,
    ran_at TEXT NOT NULL,
    ran_by TEXT,
    status TEXT NOT NULL DEFAULT 'completed',
    error TEXT
);
```

### Fields

| Field | Type | Description |
|-------|------|-------------|
| `run_id` | string | UUID v4 |
| `script_name` | string | Name of the script that was executed |
| `script_hash` | string | `sha256:<hex>` of the exact source that ran |
| `script_source` | string | **Full source snapshot** at execution time |
| `script_origin` | string | JSON object recording where the script was resolved from (see below) |
| `capabilities_used` | string | JSON array of capabilities actually invoked at runtime |
| `params` | string | JSON object of parameters (null if none). Reserved for run configs. |
| `fingerprint_before` | string | Semantic fingerprint immediately before execution |
| `fingerprint_after` | string | Semantic fingerprint immediately after execution |
| `diff_hash` | string | `sha256:<hex>` of the cell-level patch (null if no cells changed) |
| `diff_summary` | string | Compact range-level change description (see below) |
| `cells_read` | integer | Number of cell read operations |
| `cells_modified` | integer | Number of unique cells modified |
| `ops_count` | integer | Total operations queued |
| `duration_ms` | integer | Wall-clock execution time in milliseconds |
| `ran_at` | string | ISO 8601 timestamp |
| `ran_by` | string | User identifier (null if not available) |
| `status` | string | `completed`, `rolled_back`, or `error` |
| `error` | string | Error message if status is `error` (null otherwise) |

### Script origin

Every run record captures where the script was resolved from. This makes provenance auditable even when scripts move between levels.

```json
{"kind": "attached"}
{"kind": "project", "ref": ".visigrid/scripts/reconcile_q4.lua"}
{"kind": "global", "ref": "~/.config/visigrid/scripts/trim_whitespace.lua"}
{"kind": "hub", "ref": "hub:quarry/reconcile@v1.2.3"}
{"kind": "console"}
```

| `kind` | Meaning |
|--------|---------|
| `attached` | Script embedded in the .sheet file |
| `project` | Script from `.visigrid/scripts/` directory |
| `global` | Script from `~/.config/visigrid/scripts/` |
| `hub` | Script installed from VisiHub (Phase 2) |
| `console` | Ad-hoc script typed directly in the Lua console |

`ref` is the filesystem path or Hub identifier. Omitted for `attached` and `console`.

### Diff hash and diff summary

Run records capture what changed, not just that something changed.

**`diff_hash`** — `sha256:<hex>` of the canonical cell-level patch. The patch is computed as: for each cell modified, sort by `(sheet_index, row, col)`, then hash the sequence of `(sheet_index, row, col, old_value, new_value)` tuples. This enables tamper detection on the diff itself — two runs that produce identical diffs will have matching `diff_hash` values.

`null` when the script reads without writing (status `completed`, `cells_modified == 0`) or when the run errors/rolls back.

**`diff_summary`** — A compact, human-readable change description using range notation:

```
A1:A100 set (100 cells)
B2 cleared
Sheet2!C1:C50 set (50 cells), D1:D50 formulas set (50 cells)
```

Format: `[Sheet!]Range action (count)`, comma-separated. Consecutive single-cell changes to the same column are collapsed into range notation. The summary is capped at 500 characters; if it would exceed this, it ends with `... and N more regions`.

### Why snapshot full source?

If a script is modified after execution, the run record must still be independently verifiable. Storing only the hash requires the original source to be available somewhere. Storing the full source makes each run record self-contained. Lua scripts are small (typically < 10KB). The cost is negligible.

### Replay verification

```bash
vgrid sheet replay <file.sheet> --run-id <uuid>
```

Replays the run record:

1. Load the .sheet at `fingerprint_before` state (requires version history or the file at that state)
2. Execute `script_source` in sandbox
3. Compute resulting fingerprint
4. Compare against `fingerprint_after`
5. Exit 0 if match, exit 1 if mismatch

Full replay requires the file state at `fingerprint_before`. For files with version history (via VisiHub revisions), this is available. For local-only files, replay verifies the script hash and capabilities but cannot verify the before/after fingerprint transition without the prior state.

---

## 4. Console History (persistence)

### File location

```
~/.config/visigrid/console_history
```

### Format

One entry per line. Multi-line scripts are stored with `\n` escaped as literal `\\n`. Maximum 1000 entries. Oldest entries are evicted when the limit is reached.

### Behavior

- Loaded on app start into `ConsoleState.history`
- Appended to disk on each execution (same dedup rules as current in-memory history)
- Crash-safe: history is written immediately after each execution, not batched to app exit. A crash after execution but before exit loses nothing.

---

## 5. Script Library (disk storage)

### Global scripts

```
~/.config/visigrid/scripts/
  reconcile_vendors.lua
  trim_whitespace.lua
  normalize_dates.lua
```

Each `.lua` file is a script. The filename (without extension) is the script name.

### Metadata sidecar (optional)

```
~/.config/visigrid/scripts/
  reconcile_vendors.lua
  reconcile_vendors.json    # optional metadata sidecar
```

If a `.json` sidecar exists with the same basename, it provides `description`, `capabilities`, `author`, `version`, and `origin`. If no sidecar exists, `description` is extracted from the first line comment (`-- Description here`), and capabilities are inferred by static analysis.

### Project scripts

```
.visigrid/scripts/
  reconcile_q4.lua
  reconcile_q4.json         # optional
```

Same structure as global, but scoped to the project directory. Discovered when a .sheet file is opened from within a directory containing `.visigrid/`.

### Resolution order

```
attached → project → global
```

Closest to data wins. If a script with the same name exists at multiple levels, the closest one is used (shadowing).

### Script picker (Command Palette)

The Run Script picker always shows the source level and hash prefix for every script. When a script shadows another with the same name, both are listed with the shadowed entry dimmed:

```
Run: reconcile_vendors (attached · e3b0c4)
Run: reconcile_vendors (project · 7f2a91) [shadowed]
Run: trim_whitespace (global · a4c8d2)
```

Each entry shows:
- **Name** — the script name
- **Source** — `attached`, `project`, or `global`
- **Hash prefix** — first 6 hex chars of `sha256(source)`, for quick visual identity
- **Shadowed indicator** — `[shadowed]` if a closer-scoped script with the same name takes precedence

Shadowed scripts are still runnable from the picker — selecting one runs it explicitly regardless of resolution order. This lets users compare or test different versions.

### Status bar

When a .sheet has attached scripts, the status bar shows the count and a tooltip with details:

```
3 scripts attached
```

Tooltip on hover:

```
reconcile_vendors  (attached · e3b0c4)
trim_whitespace    (attached · 7f2a91)  shadows project
fill_series        (attached · a4c8d2)
```

The "shadows project" / "shadows global" annotation appears when an attached script shadows a same-named script at a lower-priority level. This gives users visibility into name collisions without requiring them to open the picker.

---

## 6. Run Configs (future, reserved)

Named tasks that bind a script to parameters and a target.

### File location

```
.visigrid/tasks/
  reconcile_q4.toml
```

### Format

```toml
name = "reconcile_q4"
script = "reconcile.lua"
description = "Q4 vendor reconciliation against bank export"

[params]
source = "bank_export.csv"
target_sheet = "Summary"
tolerance = 0.01
```

### Behavior

- Discovered alongside project scripts
- Appear in Command Palette as "Task: reconcile_q4"
- Parameters are passed to the Lua script via a global `params` table
- Each execution produces a run record with the `params` field populated

### Not implemented in Phase 1

This section reserves the schema shape. Implementation follows after run records and script library are stable.

---

## 7. Fingerprint Boundary (formal definition)

### Included in semantic fingerprint

| Data | Table | Rationale |
|------|-------|-----------|
| Cell values | `cells` | Core data |
| Cell formulas | `cells` | Computation definition |
| Cell metadata | `cell_metadata` | Semantic annotations that affect meaning |
| Iteration settings | `meta` | Affects computed values |

### Excluded from semantic fingerprint

| Data | Table/Location | Rationale |
|------|---------------|-----------|
| Formatting | `cells` (fmt_* columns) | Presentation only |
| Column widths | `col_widths` | Layout only |
| Row heights | `row_heights` | Layout only |
| Hidden rows/cols | `hidden_rows`, `hidden_cols` | View state only |
| Merged regions | `merged_regions` | Presentation only |
| Named ranges | `named_ranges` | Navigation aid (does not change cell values) |
| Hub link | `hub_link` | External reference |
| Attached scripts | `scripts` | Code, not data |
| Run records | `run_records` | History, not data |
| Style table | `meta` (style entries) | Presentation only |
| Active sheet index | `meta` | UI state |

This boundary is frozen. Changes require a fingerprint version bump (v2 → v3).

---

## 8. Hub Integration (Phase 2, reserved)

When Hub script distribution is implemented:

- Hub scripts are pulled into project or global scope (explicit `vgrid script install org/name@v1`)
- `origin` field records provenance: `"hub:org/name@v1.2.3"`
- Hub verifies `hash` matches published content
- Hub may provide trust signals (author verified, org-signed, download count)
- Hub scripts are never auto-installed or auto-updated

The `origin`, `author`, and `version` fields on script metadata exist to support this without schema changes.
