# Phase 1: Session Server (IPC)

**Status:** ✅ FROZEN (2026-02-03)
**Effort:** 2-3 weeks (actual)
**Transport:** TCP localhost + token (cross-platform, simplest to ship right)

> **Phase 1 is complete.** Server, CLI, protocol, and resource limits are implemented and tested.
> GUI controls and "unauthenticated read-only" toggle are Phase 2 polish.
> Only bugfixes allowed in this layer going forward.

---

## ⚠️ Protocol v1 Frozen

**Protocol v1 is frozen. Any changes require:**
1. `protocol_version` bump to v2
2. New golden vector files
3. Backwards compatibility analysis

Golden vectors lock the wire format:
- Server: `gpui-app/src/session_server/protocol_golden/*.jsonl`
- CLI: `crates/cli/tests/golden_vectors/*.jsonl`

Shared types: `crates/protocol/src/lib.rs` (single source of truth)

---

## Phase 1 Freeze Policy

**This layer is frozen as of 2026-02-03. The rules below govern what changes are allowed.**

### Allowed (bugfixes only)

| Category | Examples |
|----------|----------|
| Behavior behind existing fields | Fix race in writer lease renewal |
| Metrics correctness | Fix dropped event counter underflow |
| Enforcing existing limits | Tighten message size validation |
| Security hardening | Add constant-time token comparison |
| CLI behavior (no protocol change) | Add `--wait` retry loop for `apply` |
| Documentation | Clarify error code semantics |

### NOT Allowed (requires protocol v2)

| Category | Examples |
|----------|----------|
| New message types | `trace_precedents`, `undo` |
| New fields on existing messages | Adding `cell_type` to `InspectResult` |
| New optional fields | Even "harmless" additions break byte-exact tests |
| Changed semantics | Redefining what `revision` means |
| New error codes | Must stay within existing taxonomy |

### Enforcement

1. **Golden vectors are the gate.** Any PR that touches protocol types must update both server and CLI golden tests.
2. **Byte-exact tests catch drift.** If serialization changes, the test fails. Do not "fix" the golden vector — fix the code or bump protocol version.
3. **Code review checkpoint.** Any change to `crates/protocol/` or `gpui-app/src/session_server/protocol.rs` requires explicit justification.

### Phase 2 Design Constraints

The `MAX_CONNECTIONS=5` cap creates pressure for Phase 2:

- **Connection reuse**: `view --session` should hold one connection, not open/close repeatedly
- **Stats awareness**: Document "don't poll faster than 1/sec" or add `--watch` with built-in throttle
- **Multiplexing consideration**: If connection pressure grows, Phase 2 may need to multiplex operations over fewer connections

These are Phase 2 concerns. Document them here so they're not forgotten.

---

## Implementation Progress

### Completed (engine layer)

- **Revision tracking** — `Workbook.revision()` increments exactly once per successful batch. Enables optimistic concurrency control via `expected_revision`.

- **Event types** — `crates/engine/src/events.rs` defines:
  - `BatchApplied` — emitted after apply_ops (success or partial failure)
  - `CellsChanged` — cells that changed, tagged with revision
  - `RevisionChanged` — new/previous revision pair

- **Engine harness** — `crates/engine/src/harness.rs` provides:
  - `EngineHarness` wrapper with `apply_ops(ops, atomic)`
  - Event collection via `EventCollector`
  - Undo group tracking
  - Used for invariant tests without GUI dependencies

- **Invariant tests** — 16 passing tests that define the protocol contract:
  - `invariant_rollback_no_events` — atomic rollback emits only BatchApplied error
  - `invariant_rollback_no_undo_entries` — atomic rollback creates no undo group
  - `invariant_events_no_cross_revision_coalesce` — CellsChanged per revision isolated
  - `invariant_partial_nonatomic_event_semantics` — partial apply event contract
  - `invariant_revision_increments_by_one_per_batch` — no skipping, no gaps
  - `invariant_fingerprint_golden_vector_v1` — frozen encoding for v1
  - `invariant_float_canonicalization` — NaN/inf rejected, -0.0 == +0.0
  - `invariant_discovery_file_atomic` — write-to-temp + rename pattern
  - Plus 8 more covering recalc counts, revision stability, fingerprint encoding

### Completed (GUI/server layer)

- **TCP server** — `gpui-app/src/session_server/server.rs`:
  - Binds to 127.0.0.1:<random_port> (IPv4 only)
  - Non-blocking listener with shutdown signal
  - Per-connection thread spawning
  - Mode: Off | ReadOnly | Apply

- **Discovery files** — `gpui-app/src/session_server/discovery.rs`:
  - Platform-specific paths (Linux/macOS/Windows)
  - Atomic write (temp + rename)
  - 32-byte cryptographic token with constant-time comparison
  - Stale session cleanup (PID check)

- **JSONL protocol** — `gpui-app/src/session_server/protocol.rs`:
  - Message types: Hello, Welcome, ApplyOps, Inspect, Subscribe, Unsubscribe, Ping/Pong, Stats
  - Op enum: SetCellValue, SetCellFormula, ClearCell, SetNumberFormat, SetStyle
  - Error taxonomy with 11 structured error codes
  - 10MB message size limit

- **Bridge pattern** — `gpui-app/src/session_server/bridge.rs`:
  - `SessionBridgeHandle` with mpsc::Sender for cross-thread communication
  - `SessionRequest` enum routed to GUI thread
  - Oneshot response channels for request/response correlation
  - All responses include `current_revision` per spec

- **GUI integration** — `gpui-app/src/app.rs`:
  - `Spreadsheet` owns `Receiver<SessionRequest>` + `SessionServer`
  - `drain_session_requests()` called at render start
  - `handle_session_apply_ops()` applies ops via canonical mutation path:
    - Uses `batch_guard()` for single recalc + revision increment
    - Uses `set_cell_value_tracked()`/`clear_cell_tracked()` methods
    - Records history for undo (one entry per sheet with changes)
    - Marks document as modified
  - `handle_session_inspect()` queries workbook state:
    - Supports Cell, Range, and Workbook targets
    - Returns display value, raw value, and formula status
  - `start_session_server()` / `stop_session_server()` public API

- **Rate limiter** — `gpui-app/src/session_server/rate_limiter.rs`:
  - Token bucket algorithm (configurable burst/refill)
  - Per-connection rate limiting
  - Ops-based costing (apply_ops costs by op count)
  - `rate_limited` error with `retry_after_ms`

- **Event subscription** — `gpui-app/src/session_server/events.rs`:
  - Subscribe/unsubscribe to topics (currently: `cells`)
  - Bounded event queue (256 depth) with backpressure
  - Events dropped silently when queue full (best-effort delivery)
  - Revision field in every event for gap detection

- **Event coalescing** — `gpui-app/src/session_server/coalesce.rs`:
  - Converts cell sets to minimal rectangular ranges
  - O(n log n) algorithm: row bucketing → horizontal runs → vertical merge
  - Bounding box fallback when ranges exceed 2000 per sheet
  - Reduces event payload size and client redraw churn

- **Writer lease** — `gpui-app/src/session_server/server.rs`:
  - Only one connection can write at a time
  - 10-second lease, renewed on each apply_ops
  - `writer_conflict` error with `retry_after_ms` for competing clients
  - Released on disconnect

- **Connection limit** — `gpui-app/src/session_server/server.rs`:
  - `MAX_CONNECTIONS = 5` enforced at accept loop
  - Excess connections refused immediately (before handshake)
  - `connections_refused_limit` counter for stats
  - Slot freed immediately on disconnect

- **Stats endpoint** — Query server health without logs:
  - `connections_closed_parse_failures`
  - `connections_closed_oversize`
  - `writer_conflict_count`
  - `connections_refused_limit`
  - `dropped_events_total`
  - `active_connections`

- **Protocol golden vectors** — `gpui-app/src/session_server/protocol_golden/`:
  - 11 golden files locking wire format
  - Round-trip test ensures serialization stability
  - Error code coverage test ensures taxonomy completeness

- **Smoke test** — `gpui-app/src/bin/vg_session_smoke.rs`:
  - CI-ready spawn mode (`--spawn-mode`)
  - Exercises: connect, auth, apply_ops, inspect, subscribe, events
  - Verifies revision tracking, error handling, event delivery

- **Tests** — 66 session server tests + 16 invariant tests + 85 CLI tests + full gpui test suite

### Completed (CLI layer)

- **Shared protocol crate** — `crates/protocol/src/lib.rs`:
  - Canonical v1 types shared between CLI and GUI
  - `ClientMessage`, `ServerMessage` enums
  - All message types: Hello, Welcome, ApplyOps, Inspect, Stats, etc.
  - `PROTOCOL_VERSION: u32 = 1` constant

- **Session client** — `crates/cli/src/session.rs`:
  - TCP connection with bounded JSONL reads (10MB cap)
  - Request/response correlation via message ID
  - Mid-frame connection close detection
  - Unit tests for bounded read edge cases

- **Exit code registry** — `crates/cli/src/exit_codes.rs`:
  - Single source of truth for all CLI exit codes
  - Range-based organization: 0-2 universal, 3-9 diff, 10-19 AI, 20-29 session, 30-39 replay
  - `session_exit_code()` maps `SessionError` to exit codes
  - `SessionErrorOutput` for structured error output (JSON + human)

- **CLI commands** — `crates/cli/src/main.rs`:
  - `visigrid sessions` — List active sessions (table + JSON)
  - `visigrid attach` — Interactive session shell
  - `visigrid apply` — Apply ops from file/stdin (atomic + expected_revision + --wait)
  - `visigrid inspect` — Query cell/range state (table default + JSON)
  - `visigrid stats` — Session server health metrics (table + JSON)
  - `visigrid view` — Live grid snapshot (ASCII table + --follow)

- **Golden vector tests** — `crates/cli/tests/protocol_golden.rs`:
  - 17 tests including 5 byte-exact serialization tests
  - Locks wire format for both directions (GUI→CLI and CLI→GUI)
  - Uses same golden files as server-side tests

- **Tests** — 85 total CLI tests (including session + golden vector tests)

### Deferred (Phase 2 polish)

- [ ] GUI controls (Live Control panel, status bar indicator, token rotation)
- [ ] "Allow unauthenticated read-only" toggle

## Guarantees (v1)

These are the contracts clients can rely on:

| Guarantee | Implementation |
|-----------|----------------|
| Token never on disk | Discovery file has `token_hint` only (first 4 chars) |
| Deterministic ops | Fingerprint via blake3 over canonical binary encoding |
| Revision-based concurrency | `expected_revision` for optimistic locking |
| Single writer | Writer lease with 10s timeout, explicit conflicts |
| Bounded memory | 256-event queue, 10MB message limit, 2000 range cap |
| Best-effort events | Drops under backpressure; use revision gaps to detect |
| Queryable health | Stats endpoint for instant diagnostics |

## Goal

A running VisiGrid GUI exposes a local session endpoint.
External clients (CLI, agent runner, scripts) can:

1. Discover sessions
2. Apply a batch of spreadsheet operations
3. Receive structured success/error results
4. Optionally subscribe to live state (preview)

This becomes the spine for "agent builds spreadsheet with live preview."

## Non-Goals (v1)

- Remote/network access (LAN/WAN) — local only
- Full Excel formatting model — just enough style primitives
- Arbitrary GUI automation (menus, dialogs, pointer events)
- Realtime multi-writer collaboration — one writer at a time
- Precedent/dependent graph queries (deferred to v2)
- Unicode normalization (deferred to v2 with fingerprint version bump)

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                     VisiGrid GUI                             │
│                                                              │
│  ┌─────────────────────────────────────────────────────┐   │
│  │              Session Server                          │   │
│  │                                                      │   │
│  │   127.0.0.1:<random_port> (IPv4 only)               │   │
│  │   Token: <32 bytes base64>                          │   │
│  │   Mode: Off | Read-only | Apply                     │   │
│  │                                                      │   │
│  └──────────────────────┬───────────────────────────────┘   │
│                         │                                    │
│                         ▼                                    │
│  ┌─────────────────────────────────────────────────────┐   │
│  │              Workbook + Engine                       │   │
│  │              (One True Mutation Path)                │   │
│  └─────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
         ▲
         │ JSONL over TCP
         │
┌─────────────────────────────────────────────────────────────┐
│                    CLI / Agent                               │
│                                                              │
│  visigrid-cli attach                                        │
│  visigrid-cli apply --session <id> ops.jsonl                │
└─────────────────────────────────────────────────────────────┘
```

## Transport Decision

**TCP localhost + token** (not Unix sockets / named pipes)

Rationale:
- Single codepath across Linux, macOS, Windows
- Rust ecosystem has excellent async TCP support
- Token auth is well-understood pattern
- Can always add UDS later if needed

### Binding Rules

- Bind specifically to `127.0.0.1:<random_high_port>` (IPv4 only)
- Do NOT bind to `localhost` (may resolve to `::1` on some systems)
- Do NOT bind to `0.0.0.0` (LAN exposure)
- IPv6 (`::1`) support deferred to v2 — document this explicitly
- `SO_REUSEADDR` allowed (not required) for quick restart after crash

## Session Identity

On GUI start (or on "Enable Live Control"):

```rust
struct Session {
    session_id: Uuid,           // Random UUID
    token: [u8; 32],            // Random, NEVER written to disk
    port: u16,                  // Bound port
    pid: u32,                   // GUI process ID
    workbook_path: Option<PathBuf>,
    workbook_title: String,
    created_at: DateTime<Utc>,
}
```

### Discovery File

Platform-specific paths:

| Platform | Path |
|----------|------|
| Linux | `$XDG_STATE_HOME/visigrid/sessions/<id>.json` (fallback: `~/.local/state/visigrid/sessions/`) |
| macOS | `~/Library/Application Support/VisiGrid/sessions/<id>.json` |
| Windows | `%LOCALAPPDATA%\VisiGrid\sessions\<id>.json` |

```json
{
  "session_id": "a1b2c3d4-e5f6-...",
  "port": 52341,
  "pid": 1234,
  "workbook_path": "/home/user/model.sheet",
  "workbook_title": "model.sheet",
  "created_at": "2026-02-02T22:45:00Z",
  "token_hint": "a1b2..."
}
```

**Critical:** Token is NOT in the file. Only `token_hint` (first 4 chars) for
human identification. Token is displayed in GUI only (copy button) or passed
via environment variable.

### Stale Session Cleanup

CLI `sessions` command must:
1. Check if PID is still alive (platform-specific)
2. Ignore sessions older than 24h (TTL)
3. Optionally: attempt TCP connect to verify liveness

On GUI exit: delete discovery file. On crash: stale file remains (handled by CLI).

## Security Model

**Default stance:** GUI is the authority. External control is opt-in.

### Modes

| Mode | Server | Clients can |
|------|--------|-------------|
| **Off** (default) | Not running | Nothing |
| **Read-only** | Running | `inspect`, `subscribe` (token required) |
| **Apply enabled** | Running | All operations (token required) |

### Authentication

**Token required for ALL modes by default.** This prevents local malware from
scraping sheet state even in read-only mode.

Token is sent in `hello` payload:

```json
{
  "id": "1",
  "type": "hello",
  "payload": {
    "client": "visigrid-cli",
    "version": "0.4.5",
    "protocol_version": 1,
    "token": "base64-encoded-32-bytes"
  }
}
```

Connection-level auth:
- Token required for all operations (read and write)
- Wrong token → `hello_error` + close connection
- Missing token → `hello_error` + close connection

Optional GUI toggle: "Allow unauthenticated read-only" (default OFF).

### Resource Limits

| Limit | Value | Rationale |
|-------|-------|-----------|
| Max connections | 5 | Prevent resource exhaustion |
| Max ops per message | 50,000 | Reasonable batch size |
| Max ops per second | 20,000 | Prevents runaway agents |
| Max message size | 10 MB | Memory protection |
| Max line length | 10 MB | Before buffering entire line |
| Max cells affected per style op | 100,000 | See style storage note below |

Note: Rate limiting is **ops-based**, not request-based. A single message
with 50k ops counts as 50k against the rate limit.

### Style Storage Constraint

**Style operations MUST be represented internally as range style spans
(not materialized per-cell) to keep time/memory bounded.**

If engine stores style per-cell, the 100k cap is still dangerous. If style
is stored as range overlays, cap can be raised. This constraint is why
the default is 100k not 1M.

## Protocol: JSONL over TCP

### Protocol Version

Current version: **1**

Versioning rules:
- Client sends `protocol_version` in `hello`
- Server returns `protocol_version` and `capabilities` in `hello_ok`
- If client version > server version, server returns `hello_error` with supported version
- Protocol version changes when wire format changes incompatibly

Future: `engine_version` for determinism guarantees across VisiGrid releases.

### Message Framing

Newline-delimited JSON (JSONL):
- Each message is a single line terminated by `\n`
- Input must be valid UTF-8
- Each line must parse as exactly one JSON object
- Max line length enforced **before** buffering (10 MB)

On parse failure:
- Respond with `bad_request` error
- Keep connection open (unless repeated failures)
- After 3 consecutive parse failures, close connection

```json
{
  "id": "client-generated-id",
  "type": "message_type",
  "payload": { ... }
}
```

## Error Taxonomy

All errors use a fixed code from this taxonomy:

| Code | Description |
|------|-------------|
| `auth_failed` | Invalid or missing token |
| `protocol_mismatch` | Unsupported protocol version |
| `rate_limited` | Ops/sec limit exceeded |
| `message_too_large` | Message exceeds 10 MB |
| `op_limit_exceeded` | Too many ops in single message |
| `cells_limit_exceeded` | Style op affects too many cells |
| `revision_mismatch` | expected_revision doesn't match current |
| `invalid_op` | Malformed operation |
| `formula_parse_error` | Formula syntax error |
| `eval_error` | Formula evaluation error (circular ref, etc.) |
| `out_of_bounds` | Cell reference outside valid range |
| `sheet_not_found` | Sheet index doesn't exist |
| `bad_request` | Malformed JSON or unknown message type |

All `apply_ops_error` responses include:
- `code` — error code from taxonomy
- `message` — human-readable description
- `op_index` — which op failed (0-indexed)
- `location` — optional `{sheet, cell}` or `{sheet, range}`
- `current_revision` — always included for recovery

## Core Message Types

### 1. `hello` — Handshake + Auth

**Client → Server:**
```json
{
  "id": "1",
  "type": "hello",
  "payload": {
    "client": "visigrid-cli",
    "version": "0.4.5",
    "protocol_version": 1,
    "token": "dGhpcyBpcyBhIDMyIGJ5dGUgdG9rZW4uLi4="
  }
}
```

**Server → Client (success):**
```json
{
  "id": "1",
  "type": "hello_ok",
  "payload": {
    "session_id": "a1b2c3d4-...",
    "protocol_version": 1,
    "capabilities": ["apply_ops", "inspect", "subscribe"],
    "workbook": {
      "title": "model.sheet",
      "sheets": 3,
      "revision": 42
    },
    "limits": {
      "max_ops_per_message": 50000,
      "max_cells_per_style_op": 100000
    }
  }
}
```

**Server → Client (auth failure):**
```json
{
  "id": "1",
  "type": "hello_error",
  "payload": {
    "code": "auth_failed",
    "message": "Invalid or missing token",
    "supported_protocol_versions": [1]
  }
}
```
Connection closed after `hello_error`.

### 2. `apply_ops` — Batch mutation

**Client → Server:**
```json
{
  "id": "2",
  "type": "apply_ops",
  "payload": {
    "batch_name": "agent_step_12",
    "atomic": true,
    "expected_revision": 42,
    "ops": [
      { "op": "set_cell_value", "sheet": 0, "cell": "A1", "value": "Revenue" },
      { "op": "set_cell_formula", "sheet": 0, "cell": "B1", "formula": "=SUM(B2:B100)" },
      { "op": "set_style", "sheet": 0, "range": "A1:B1", "bold": true },
      { "op": "set_number_format", "sheet": 0, "range": "B2:B100", "format": "currency_usd" }
    ]
  }
}
```

**`expected_revision`** (optional but recommended):
- If provided and doesn't match current revision, returns `revision_mismatch` error
- Prevents stale writer from overwriting newer state
- Client can re-inspect and retry

**Server → Client (success):**
```json
{
  "id": "2",
  "type": "apply_ops_ok",
  "payload": {
    "applied": 4,
    "undo_group_id": "u_9c1",
    "fingerprint": "v1:4:a3f2b1c4d5e6f7a8b9c0d1e2f3a4b5c6",
    "recalc_ms": 12,
    "revision": 43
  }
}
```

**Server → Client (error):**
```json
{
  "id": "2",
  "type": "apply_ops_error",
  "payload": {
    "code": "formula_parse_error",
    "message": "Unexpected token ')'",
    "op_index": 1,
    "location": { "sheet": 0, "cell": "B1" },
    "current_revision": 42
  }
}
```

### 3. `inspect` — Query cell state (v1-safe fields only)

**Client → Server:**
```json
{
  "id": "3",
  "type": "inspect",
  "payload": {
    "sheet": 0,
    "range": "B1"
  }
}
```

**Server → Client:**
```json
{
  "id": "3",
  "type": "inspect_ok",
  "payload": {
    "revision": 43,
    "cells": [
      {
        "cell": "B1",
        "value": 1234567,
        "display": "$1,234,567.00",
        "formula": "=SUM(B2:B100)",
        "format": "currency_usd",
        "style": { "bold": true },
        "error": null
      }
    ]
  }
}
```

**v1 fields only:** value, display, formula, format, style, error.
**Deferred to v2:** `precedents`, `dependents` (separate `trace_precedents` / `trace_dependents` messages).

### 4. `subscribe` — Live preview events

**Client → Server:**
```json
{
  "id": "4",
  "type": "subscribe",
  "payload": {
    "topics": ["cells_changed", "revision_changed", "batch_applied"],
    "throttle_ms": 50
  }
}
```

**Server → Client:**
```json
{ "id": "4", "type": "subscribe_ok", "payload": { "revision": 43 } }
```

### Event Ordering

For each successful `apply_ops`, events are emitted in this order:

1. `batch_applied` — batch metadata
2. `cells_changed` — affected ranges (coalesced)
3. `revision_changed` — new revision number

**Every event includes `revision` field.** Clients can use this to order
events and detect missed events.

```json
{ "type": "event", "payload": { "topic": "batch_applied", "revision": 44, "batch_name": "agent_step_12", "applied": 4, "recalc_ms": 12 } }
{ "type": "event", "payload": { "topic": "cells_changed", "revision": 44, "ranges": ["A1:B1", "B2:B100"] } }
{ "type": "event", "payload": { "topic": "revision_changed", "revision": 44 } }
```

**v1 topics:**
- `cells_changed` — coalesced ranges of changed cells
- `revision_changed` — workbook revision number changed
- `batch_applied` — external batch was applied

**Deferred:** `selection` (agents don't need it; humans watching is fluff for v1)

## Operations (Minimal Set)

Internal `Op` enum shared by GUI actions, replay engine, and session server.

### The One True Mutation Path

**Critical architectural rule:**

> GUI action handlers MUST emit Ops (or call the same apply path), not mutate
> workbook directly. This prevents "two engines" bugs where GUI and session
> server behave differently.

```rust
// CORRECT: GUI uses same path as session server
fn on_user_types_value(&mut self, cell: CellRef, value: String, cx: &mut Context<Self>) {
    let op = Op::SetCellValue { sheet: self.active_sheet, cell, value };
    self.apply_ops(vec![op], true, None, cx);  // None = no expected_revision check
}

// WRONG: GUI mutates directly, session server uses different path
fn on_user_types_value(&mut self, cell: CellRef, value: String, cx: &mut Context<Self>) {
    self.workbook.sheet_mut(0).set_value(cell, &value);  // NO!
}
```

### Cell Reference Format

Wire protocol accepts **both** A1 strings and numeric coordinates:

```json
// A1 notation (ergonomic)
{ "op": "set_cell_value", "sheet": 0, "cell": "B2", "value": "Hello" }

// Numeric coordinates (unambiguous)
{ "op": "set_cell_value", "sheet": 0, "row": 1, "col": 1, "value": "Hello" }
```

Server normalizes to numeric coordinates immediately on parse.
A1 parsing rules:
- Column letters uppercase (a1 → A1)
- No `$` in A1 (absolute refs stripped)
- Ranges: "A1:C10" or `{"start_row": 0, "start_col": 0, "end_row": 9, "end_col": 2}`

### MVP Ops (v1)

| Op | Fields | Description |
|----|--------|-------------|
| `set_cell_value` | sheet, cell/row+col, value | Set cell to literal |
| `set_cell_formula` | sheet, cell/row+col, formula | Set cell to formula |
| `clear_range` | sheet, range | Clear cells |
| `set_number_format` | sheet, range, format, decimals? | Number format |
| `set_style` | sheet, range, bold?, italic?, underline?, bg?, fg? | Text style |

### Deferred Ops (v2)

| Op | Fields | Description |
|----|--------|-------------|
| `set_border` | sheet, range, style | Cell borders |
| `add_sheet` | name | Add new sheet |
| `rename_sheet` | sheet, name | Rename sheet |
| `set_column_width` | sheet, col, width | Column sizing |
| `set_row_height` | sheet, row, height | Row sizing |

## Critical Semantics

### Unified Atomic Apply Flow

**This is the single canonical implementation.** Both success and failure paths
execute inside BatchGuard scope so recalc/render only occur once.

```rust
fn apply_ops(
    &mut self,
    ops: Vec<Op>,
    atomic: bool,
    expected_revision: Option<u64>,
    cx: &mut Context<Self>,
) -> Result<ApplyResult, ApplyError> {
    // 1. Check revision if provided
    let current_rev = self.workbook.read(cx).revision();
    if let Some(expected) = expected_revision {
        if expected != current_rev {
            return Err(ApplyError {
                code: "revision_mismatch",
                message: format!("Expected revision {}, current is {}", expected, current_rev),
                op_index: None,
                location: None,
                current_revision: current_rev,
            });
        }
    }

    // 2. Begin undo group BEFORE batch guard
    self.history.begin_group("External batch");

    // 3. Enter batch guard (defers recalc + change notifications)
    let result = self.workbook.update(cx, |wb, _| {
        let mut guard = wb.batch_guard();

        for (idx, op) in ops.iter().enumerate() {
            match apply_single_op(&mut guard, op) {
                Ok(change) => {
                    // Record for undo (still inside batch guard)
                    self.history.record_change(change);
                }
                Err(e) => {
                    if atomic {
                        // Rollback INSIDE batch guard scope
                        // abort_group applies inverse ops but recalc is still deferred
                        self.history.abort_group(&mut guard);
                        return Err(ApplyError {
                            code: e.code,
                            message: e.message,
                            op_index: Some(idx),
                            location: e.location,
                            current_revision: current_rev,
                        });
                    }
                    // Non-atomic: commit partial, return applied count
                    self.history.end_group();
                    return Ok(ApplyResult { applied: idx, ... });
                }
            }
        }

        // 4. Commit undo group
        self.history.end_group();

        Ok(ApplyResult {
            applied: ops.len(),
            fingerprint: compute_fingerprint(&ops),
            revision: wb.increment_revision(),
            ...
        })
    });  // BatchGuard drops here → single recalc

    // 5. Single render notification
    cx.notify();

    result
}
```

**Key invariants:**
- ✅ Single recalc (BatchGuard defers until drop)
- ✅ Rollback without engine transaction (undo group abort)
- ✅ No double-recording history
- ✅ No UI flicker mid-batch
- ✅ Rollback also inside BatchGuard (no double recalc on failure)

### Deterministic Fingerprint

**Fingerprint canonicalization MUST NOT rely on JSON object key ordering.**
JSON key order is not guaranteed by serde_json. Use stable binary encoding.

```rust
/// Binary canonical encoding for fingerprint.
/// Fields in fixed order, length-prefixed strings.
fn op_canonical_bytes(op: &Op) -> Vec<u8> {
    let mut buf = Vec::new();

    // Op tag (1 byte)
    buf.push(op.tag());

    // Sheet (4 bytes, little-endian)
    buf.extend_from_slice(&op.sheet().to_le_bytes());

    // Row (4 bytes)
    buf.extend_from_slice(&op.row().to_le_bytes());

    // Col (4 bytes)
    buf.extend_from_slice(&op.col().to_le_bytes());

    // Payload: length-prefixed UTF-8 bytes
    match op {
        Op::SetCellValue { value, .. } => {
            buf.extend_from_slice(&(value.len() as u32).to_le_bytes());
            buf.extend_from_slice(value.as_bytes());
        }
        Op::SetCellFormula { formula, .. } => {
            buf.extend_from_slice(&(formula.len() as u32).to_le_bytes());
            buf.extend_from_slice(formula.as_bytes());
        }
        Op::SetStyle { bold, italic, underline, bg, fg, .. } => {
            // Fixed-order boolean flags + optional color bytes
            buf.push(flags_byte(bold, italic, underline));
            write_optional_color(&mut buf, bg);
            write_optional_color(&mut buf, fg);
        }
        // ... other ops
    }

    buf
}

fn compute_fingerprint(ops: &[Op]) -> String {
    let mut hasher = blake3::Hasher::new();
    for op in ops {
        hasher.update(&op_canonical_bytes(op));
    }
    let hash = hasher.finalize();
    format!("v1:{}:{}", ops.len(), hash.to_hex())
}
```

**Canonicalization rules:**
- Cell references: normalized to numeric (row, col) internally
- Numbers: serialized as raw bytes (f64 little-endian)
- Formulas: raw UTF-8 bytes as received (no rewriting, no normalization)
- Strings: raw UTF-8 bytes as received
- **No Unicode normalization in v1** — hash exact bytes

Future v2: If Unicode normalization added, fingerprint version bumps to `v2`.

## GUI Integration

### Live Control Panel

```
┌─────────────────────────────────────────────────────────┐
│ Live Control                                    [X] │
├─────────────────────────────────────────────────────────┤
│ Status:  ● Apply Enabled                      [▼]  │
│                                                     │
│ Endpoint: 127.0.0.1:52341                   [Copy] │
│ Token:    ●●●●●●●●●●●●                      [Copy] │
│ Session:  a1b2c3d4                                 │
│                                                     │
│ Connected: 1 client                                │
│ Last batch: agent_step_12 (4 ops, 12ms)           │
│                                                     │
│ [Rotate Token]  [Disconnect All]                   │
│                                                     │
│ [ ] Allow unauthenticated read-only (not recommended) │
└─────────────────────────────────────────────────────────┘
```

### Status Bar Integration

```
┌─────────────────────────────────────────────────────────────────┐
│ [● LIVE]  model.sheet                             Row 1, Col A  │
└─────────────────────────────────────────────────────────────────┘
```

- Green dot = Apply enabled, connected
- Yellow dot = Read-only
- No dot = Off

### Toast on External Changes

When ops applied externally:

```
┌────────────────────────────────────────┐
│ External changes applied (undo: ⌘Z)   │
└────────────────────────────────────────┘
```

Disappears after 3s.

## Files to Create

| Path | Purpose |
|------|---------|
| `gpui-app/src/session/mod.rs` | Module root |
| `gpui-app/src/session/server.rs` | TCP listener, connection handling |
| `gpui-app/src/session/protocol.rs` | Message types, serialization, framing |
| `gpui-app/src/session/auth.rs` | Token generation, validation |
| `gpui-app/src/session/discovery.rs` | Session file write/cleanup (platform paths) |
| `gpui-app/src/session/ops.rs` | Op enum, canonical bytes, apply logic |
| `gpui-app/src/session/limits.rs` | Rate limiting, connection limits |
| `gpui-app/src/session/errors.rs` | Error taxonomy, structured error types |

## Dev Test Harness

**Critical:** Do not ship untested. Create a minimal test client in-repo.

```
gpui-app/src/session/test_client.rs   # Or separate binary
```

Test client capabilities:
- Connect to session by ID
- Send `hello` with token
- Send `apply_ops` batch from JSON file
- Print responses
- Basic REPL for manual testing

```bash
# Dev testing (not shipped to users)
cargo run --bin vg-session-test -- --session a1b2c3 --ops test_ops.jsonl
```

Integration tests:
- Start GUI headless (or with test window)
- Enable session server
- Connect test client
- Apply ops, verify state
- Test error cases:
  - Bad token → `auth_failed`
  - Bad formula → `formula_parse_error` with `op_index`
  - Atomic rollback → verify all ops rolled back
  - `expected_revision` mismatch → `revision_mismatch`
  - Rate limit exceeded → `rate_limited`

## Build Milestones

### Milestone 1: Server Skeleton (3-5 days)

- [ ] TCP listener on 127.0.0.1:<random> (IPv4 only)
- [ ] Token generation (32 bytes)
- [ ] `hello` handshake with protocol version
- [ ] Connection-level auth (token required)
- [ ] Discovery file write/cleanup (platform paths)
- [ ] Stale session detection (PID check)
- [ ] GUI toggle: Off / Read-only / Apply
- [ ] Status bar indicator
- [ ] Message framing (JSONL with size limits)
- [ ] Error taxonomy types
- [ ] Dev test client (basic connect + hello)

### Milestone 2: Apply Ops (5-7 days)

- [ ] Define `Op` enum with canonical bytes
- [ ] Cell reference parsing (A1 + numeric)
- [ ] Wire ops to workbook mutations via One True Path
- [ ] Unified apply flow (undo inside batch guard)
- [ ] Atomic rollback via undo group abort
- [ ] `expected_revision` check
- [ ] Structured errors with `op_index` + `current_revision`
- [ ] `apply_ops` message handling
- [ ] Fingerprint computation (blake3, binary canonical)
- [ ] Dev test client (apply_ops testing)

### Milestone 3: Query + Events (3-5 days)

- [ ] `inspect` message (v1 fields only)
- [ ] `subscribe` message
- [ ] Event ordering (batch_applied → cells_changed → revision_changed)
- [ ] Mandatory `revision` in every event
- [ ] Throttle/coalesce event stream
- [ ] Workbook revision tracking

### Milestone 4: Safety Polish (2-3 days)

- [ ] Connection limit (max 5)
- [ ] Ops-per-second rate limiting
- [ ] Ops-per-message limit
- [ ] Cells-per-style-op limit
- [ ] Token rotation (GUI button)
- [ ] Graceful disconnect on GUI close
- [ ] Error handling for malformed messages
- [ ] "Allow unauthenticated read-only" toggle (default OFF)

## CLI Commands (COMPLETE)

All session CLI commands are implemented. See `crates/cli/src/main.rs`.

```bash
# List running sessions (with stale PID detection)
visigrid sessions                    # Table format
visigrid sessions --json             # JSON format

# Apply ops to running session
visigrid apply ops.jsonl                              # Auto-detect session
visigrid apply --session a1b2c3 ops.jsonl             # Explicit session
visigrid apply --expected-revision 42 ops.jsonl       # Optimistic concurrency
visigrid apply --atomic=false ops.jsonl               # Partial apply mode
cat ops.jsonl | visigrid apply -                      # Stdin
visigrid apply --wait ops.jsonl                       # Retry on writer conflict
visigrid apply --wait --wait-timeout 60 ops.jsonl     # Custom timeout (default 30s)

# Query cell/range state
visigrid inspect A1                  # Single cell (table)
visigrid inspect A1:C10              # Range (table)
visigrid inspect A1 --json           # JSON format

# Session server health metrics
visigrid stats                       # Table format
visigrid stats --json                # JSON format

# View live grid (read-only snapshot)
visigrid view                        # Default: A1:J20
visigrid view --range A1:K30         # Custom range
visigrid view --follow               # Auto-refresh on changes
visigrid view --width 15             # Custom column width

# Interactive shell
visigrid attach                      # REPL for exploration
```

### Exit Codes

Session commands use exit codes 20-29 for scripting:

| Code | Constant | Description |
|------|----------|-------------|
| 20 | `EXIT_SESSION_CONNECT` | Cannot connect (no server, refused) |
| 21 | `EXIT_SESSION_PROTOCOL` | Protocol error (bad framing, version) |
| 22 | `EXIT_SESSION_AUTH` | Authentication failed (bad token) |
| 23 | `EXIT_SESSION_CONFLICT` | Writer conflict or revision mismatch |
| 24 | `EXIT_SESSION_PARTIAL` | Partial apply (some ops rejected) |
| 25 | `EXIT_SESSION_INPUT` | Invalid input (bad op, bad reference) |
| 26 | `EXIT_SESSION_TIMEOUT` | Operation timed out |

### Writer Conflict Retry (`--wait`)

When `apply` encounters a `writer_conflict` error:
- Without `--wait`: fails immediately with exit 23
- With `--wait`: retries with adaptive backoff until success or timeout
- `--wait-timeout N`: custom timeout in seconds (default 30)
- On timeout: exit 23 (`EXIT_SESSION_CONFLICT`)

**Adaptive backoff:**
- Uses `retry_after_ms` from server (clamped to 50-2000ms)
- Adds ±10% jitter to prevent thundering herd
- Reuses same connection (saves connection slots)
- Reconnects automatically if connection drops

**Safety guard:**
```
--wait requires --atomic or --expected-revision
```
Without idempotency protection, retrying can cause double-apply. Exit 2 (usage error) if neither is provided.

This makes scripts reliable without manual retry loops or hidden correctness bugs.

### Live Grid View (`view`)

Read-only grid snapshot with optional auto-refresh:
- Displays range as ASCII table with row/column headers
- `--follow`: polls revision every 500ms, refreshes on change
- Single connection held for duration (respects connection cap)
- Truncates cells to `--width` characters (default 12)

```
Session: a1b2c3d4  Sheet: 0  Range: A1:J20  Revision: 43
────────────────────────────────────────────────────────────
            A           B           C           D
─────────────────────────────────────────────────────────
    1    Revenue     Q1 2024     Q2 2024     Q3 2024
    2    Product A     10000       12000       15000
    3    Product B      8000        9500       11000
```

### Bounded JSONL Reads

CLI enforces 10MB message limit before buffering:
- Prevents memory exhaustion from malicious/buggy servers
- Detects mid-frame connection close (no newline)
- Returns `EXIT_SESSION_PROTOCOL` (21) on violation

### Deferred

```bash
# Watch file and re-apply on save (not implemented)
visigrid watch ops.jsonl --session a1b2c3
```

## The "Prove It" Demo

1. Start GUI, enable Apply mode
2. Run agent that emits ops to build a revenue model
3. GUI updates live as each batch applies
4. Introduce deliberate bad formula → receive structured error with `op_index` + `current_revision` → agent fixes → success
5. Attempt batch with error at op 3 → verify atomic rollback (ops 0-2 also rolled back, single recalc)
6. Attempt batch with wrong `expected_revision` → receive `revision_mismatch` → agent re-inspects and retries
7. Save `.sheet`
8. Run `visigrid-cli replay --verify` → fingerprint matches

This demo earns the "compiler loop" marketing.

---

## Implementation Invariants

These invariants MUST be enforced in code and verified by tests BEFORE feature work begins.

### 1. Silent Mutation Mode for Rollback

`abort_group()` must run in "silent mutation" mode:

```rust
impl History {
    /// Rollback all changes in current group.
    /// MUST NOT: create new undo entries, emit events, notify UI.
    /// MUST: operate under same BatchGuard (caller's responsibility).
    fn abort_group(&mut self, guard: &mut BatchGuard) {
        self.recording_enabled = false;  // Disable history recording
        for change in self.pending_group.drain(..).rev() {
            change.apply_inverse(guard);  // Apply inverse under same guard
        }
        self.recording_enabled = true;
        // No cx.notify() here — caller does it after guard drops
    }
}
```

If your undo system operates at a higher layer (outside engine guard), refactor it first.

### 2. Float Canonicalization

For numeric payloads in fingerprint:

```rust
fn canonicalize_float(f: f64) -> [u8; 8] {
    if f.is_nan() {
        // Canonical NaN: all NaN values → single bit pattern
        return f64::NAN.to_le_bytes();  // Or reject at protocol level
    }
    if f == 0.0 {
        // Canonical zero: -0.0 and +0.0 → same bytes
        return 0.0_f64.to_le_bytes();
    }
    f.to_le_bytes()
}
```

**Protocol-level rule:** Reject `NaN`, `+inf`, `-inf` in `apply_ops`. Return `invalid_op` error.
Rationale: Financial data should never contain these values.

### 3. Revision Increment Semantics

Revision is monotonic per workbook session:

| Event | Revision change |
|-------|-----------------|
| Successful `apply_ops` (even no-op batch) | +1 |
| GUI edit via One True Path | +1 |
| Failed `apply_ops` (atomic rollback) | No change |
| `revision_mismatch` rejection | No change |
| `inspect`, `subscribe` | No change |
| File save | No change (revision is in-memory session state) |
| File load | Reset to 0 (or loaded value) |

**Invariant:** Any mutation via One True Path increments revision exactly once per batch/group.

### 4. Event Coalescing Boundaries

**Invariant:** `cells_changed` ranges MUST only represent changes for the specific revision attached to the event.

If agent applies 3 batches quickly (revisions 45, 46, 47), subscriber receives 3 separate `cells_changed` events, not one merged event.

Coalescing is allowed **within** a single revision (multiple ranges → fewer ranges), but NEVER **across** revisions.

```rust
// CORRECT: Each batch emits its own events
// Batch 1 (rev 45): cells_changed { revision: 45, ranges: ["A1:A10"] }
// Batch 2 (rev 46): cells_changed { revision: 46, ranges: ["B1:B10"] }

// WRONG: Merged across revisions
// cells_changed { revision: 47, ranges: ["A1:A10", "B1:B10"] }  // NO!
```

### 5. Rate Limiting: Token Bucket

```rust
struct RateLimiter {
    tokens: u32,
    max_tokens: u32,      // Burst capacity: 40,000
    refill_rate: u32,     // Per second: 20,000
    last_refill: Instant,
}

impl RateLimiter {
    fn try_consume(&mut self, ops: u32) -> Result<(), RateLimitedError> {
        self.refill();
        if ops > self.tokens {
            return Err(RateLimitedError { retry_after_ms: ... });
        }
        self.tokens -= ops;
        Ok(())
    }
}
```

- Token bucket: 20k ops/sec refill, 40k burst capacity
- Per-connection
- `rate_limited` error includes `retry_after_ms` hint

### 6. Discovery File: Permissions + Atomic Write

```rust
fn write_discovery_file(session: &Session) -> io::Result<()> {
    let dir = discovery_dir();
    fs::create_dir_all(&dir)?;

    let final_path = dir.join(format!("{}.json", session.session_id));
    let temp_path = dir.join(format!("{}.json.tmp", session.session_id));

    // Write to temp file
    let json = serde_json::to_string_pretty(&session.to_discovery())?;
    fs::write(&temp_path, &json)?;

    // Set restrictive permissions (best effort)
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        fs::set_permissions(&temp_path, fs::Permissions::from_mode(0o600))?;
    }

    // Atomic rename
    fs::rename(&temp_path, &final_path)?;

    Ok(())
}
```

### 7. Recalc Count Invariant

**Invariant:** Every `apply_ops` call results in exactly ONE recalculation, regardless of:
- Number of ops in batch
- Success or failure (atomic rollback)
- Number of cells affected

Test by instrumenting `recalc_dirty_set()` with a counter.

---

## Invariant Tests (Write First)

Before implementing features, write these 10 tests:

```rust
#[cfg(test)]
mod invariant_tests {
    // 1. Rollback under batch guard doesn't emit events
    #[test]
    fn rollback_no_events() { ... }

    // 2. Rollback doesn't create undo entries
    #[test]
    fn rollback_no_undo_entries() { ... }

    // 3. Recalc count = 1 on success
    #[test]
    fn single_recalc_on_success() { ... }

    // 4. Recalc count = 1 on failure (atomic rollback)
    #[test]
    fn single_recalc_on_failure() { ... }

    // 5. Revision increments exactly once per successful batch
    #[test]
    fn revision_increments_once_per_batch() { ... }

    // 6. Revision doesn't increment on rejected batch
    #[test]
    fn revision_stable_on_rejection() { ... }

    // 7. cells_changed never spans multiple revisions
    #[test]
    fn events_per_revision_boundary() { ... }

    // 8. Fingerprint stable across platforms (same ops → same hash)
    #[test]
    fn fingerprint_deterministic() { ... }

    // 9. Float canonicalization (NaN/inf rejected, -0.0 == +0.0)
    #[test]
    fn float_canonicalization() { ... }

    // 10. Discovery file atomic + readable after write
    #[test]
    fn discovery_file_atomic() { ... }
}
```

Run these in CI matrix (Linux, macOS, Windows) to catch platform-specific nondeterminism.

---

## Observability

Instrument with structured logging. One log line per batch:

```json
{
  "event": "batch_applied",
  "connection_id": "conn_1",
  "client": "visigrid-cli",
  "client_version": "0.4.5",
  "protocol_version": 1,
  "batch_name": "agent_step_12",
  "ops_count": 4,
  "applied": 4,
  "recalc_ms": 12,
  "revision_before": 42,
  "revision_after": 43,
  "fingerprint": "v1:4:a3f2b1c4..."
}
```

On error:

```json
{
  "event": "batch_rejected",
  "connection_id": "conn_1",
  "client": "visigrid-cli",
  "error_code": "formula_parse_error",
  "op_index": 1,
  "revision": 42,
  "message": "Unexpected token ')'"
}
```

This pays for itself immediately when debugging agent issues.
