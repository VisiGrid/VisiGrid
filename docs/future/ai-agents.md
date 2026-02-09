# AI Agents: VisiGrid as a Headless Spreadsheet Engine

**Status:** Planned (post-HN launch)

## Overview

VisiGrid provides the "Compiler Loop" for spreadsheet logic that LLMs lack.
Instead of hallucinating math in context windows, agents interact with a
deterministic engine that gives immediate feedback on formulas.

**The core insight:** Agents don't need to parse a 2D grid. They need to
interact with a Schema + DAG via structured operations.

## Why Agents Fail at Spreadsheets Today

| Problem | Excel/Sheets | VisiGrid |
|---------|--------------|----------|
| Token waste | Agent parses spatial layout, loses accuracy | Agent reads logic paths only |
| Error feedback | Execute → traceback → retry | Immediate structured error |
| Auditability | Black-box script or opaque API calls | Lua provenance trail |
| Network latency | Google Sheets API rate limits | Local-first, zero latency |

## Implementation Phases

| Phase | Name | Effort | Deliverable |
|-------|------|--------|-------------|
| 0 | HN Launch | Now | CLI diff/calc — no agent mention |
| 1 | [Session Server](phase-1-session-server.md) | 2-3 weeks | TCP server + structured errors + live preview |
| 2 | [Agent Kit](phase-2-agent-kit.md) | 3-5 days | MCP tools + CLAUDE.md + demos |

**Note:** We skip the "file watcher" approach entirely. Going straight to IPC
means building the right foundation once, with proper batching semantics,
structured errors, and a real protocol.

## Phase Summary

### Phase 1: Session Server (IPC)

**The platform move.** GUI exposes a local TCP endpoint. External clients
(CLI, agents, scripts) can:

- Discover running sessions
- Apply batches of operations atomically
- Receive op-indexed structured errors
- Subscribe to live state events

**Key decisions:**

- **Transport:** TCP localhost + token (cross-platform, simple)
- **Protocol:** JSONL over single connection (debuggable)
- **Security:** Off / Read-only / Apply modes, token auth required
- **Semantics:** Atomic batches, single recalc, single undo group

**This is what makes the "compiler loop" claim defensible.**

See: [phase-1-session-server.md](phase-1-session-server.md)

### Phase 2: Agent Kit

**The launch.** Package VisiGrid as a tool for LLMs.

- MCP tool definition (formal schema)
- CLAUDE.md snippet for Claude Code users
- Example prompts + guardrails
- Demo scripts + video

See: [phase-2-agent-kit.md](phase-2-agent-kit.md)

## Architecture: The Compiler Loop

```
┌─────────────────────────────────────────────────────────────┐
│                        AI AGENT                              │
│  1. Generate operations                                      │
│  2. Send batch to VisiGrid session                          │
│  3. Read result or structured error                         │
│  4. Self-correct if needed                                  │
│  5. Repeat                                                   │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       │ JSONL over TCP (127.0.0.1)
                       │ Authorization: Bearer <token>
                       ▼
┌─────────────────────────────────────────────────────────────┐
│                VisiGrid GUI (Session Server)                 │
│                                                              │
│  - Batch ops → single recalc → single undo group           │
│  - Structured errors with op_index                          │
│  - Deterministic fingerprint                                │
│  - Live preview events                                       │
│                                                              │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       │ apply_ops_ok / apply_ops_error
                       ▼
┌─────────────────────────────────────────────────────────────┐
│                        AI AGENT                              │
│  Error? → Read op_index + suggestion → Fix → Retry          │
│  Success? → Continue to next batch                          │
└─────────────────────────────────────────────────────────────┘
```

## Protocol Overview

### Message Format (JSONL)

```json
{"id": "1", "type": "hello", "payload": {"client": "vgrid", "version": "0.4.5"}}
{"id": "2", "type": "apply_ops", "payload": {"atomic": true, "ops": [...]}}
{"id": "3", "type": "inspect", "payload": {"sheet": 0, "range": "B1"}}
{"id": "4", "type": "subscribe", "payload": {"topics": ["cells_changed"]}}
```

### Structured Error Response

```json
{
  "id": "2",
  "type": "apply_ops_error",
  "payload": {
    "error": "FormulaParseError",
    "message": "Unexpected token ')'",
    "op_index": 1,
    "location": {"sheet": 0, "cell": "B1"},
    "suggestion": "Check for unmatched parenthesis"
  }
}
```

Op-indexed errors let agents know exactly which command failed and why.

## MVP Operations

| Op | Fields | Description |
|----|--------|-------------|
| `set_cell_value` | sheet, cell, value | Set cell to literal |
| `set_cell_formula` | sheet, cell, formula | Set cell to formula |
| `clear_range` | sheet, range | Clear cells |
| `set_number_format` | sheet, range, format | Number format |
| `set_style` | sheet, range, bold?, italic?, bg?, fg? | Text style |

Minimal surface. No layout ops in v1. No full Excel formatting.

## Security Model

**Default stance:** GUI is the authority. External control is opt-in.

| Mode | Server | Clients can |
|------|--------|-------------|
| Off (default) | Not running | Nothing |
| Read-only | Running | Query state only |
| Apply enabled | Running | Query + mutate |

**Hard rules:**
- Bind to 127.0.0.1 only (no LAN exposure)
- Require `Authorization: Bearer <token>` on every request
- Token displayed in GUI only (never written to disk)
- Rate limit 100 req/sec, 10MB message cap

## Competitive Position

| Competitor | Limitation | VisiGrid Advantage |
|------------|------------|-------------------|
| Excel + VBA | Binary format, slow API, no CLI | Native protocol, Lua provenance |
| Google Sheets | Network latency, rate limits | Local-first, zero latency |
| Python/Pandas | Black-box scripts | Verifiable fingerprints |
| DuckDB | Can't hand to CFO for audit | Spreadsheet explainability |

**The one-liner:** Database performance, spreadsheet explainability.

## The Defensible Claim

When someone asks "Can agents really work better this way?":

> "We're not claiming AI is smarter. We're claiming the feedback loop is
> tighter. Instead of an agent writing a black-box Python script, it
> interacts with VisiGrid as a reactive engine. The agent sends operations,
> VisiGrid deterministically executes them, and errors come back with
> op-indexed locations — like a compiler."

## Timeline

```
Week 0-4:   HN launch (diff/calc focus, no agent mention)
Week 4-7:   Phase 1 (Session Server) — 2-3 weeks
Week 7-8:   Phase 2 (Agent Kit) — 3-5 days
Week 8+:    "VisiGrid for Agents" launch
```

## The "Prove It" Demo

1. Start GUI, enable Apply mode
2. Run agent that emits ops to build a revenue model
3. GUI updates live as each batch applies
4. Introduce deliberate bad formula → receive structured error → agent fixes → success
5. Save `.sheet`
6. Run `vgrid replay --verify` → fingerprint matches

This demo earns the marketing.

## Open Questions

1. **Pairing flow?** Optional: GUI shows short code, client supplies it to
   mint token. Nice for human-in-the-loop approval, but not required for v1.

2. **Sandbox mode?** "Apply to scratch copy" is nice-to-have for extra safety.
   Defer unless users demand it.

3. **Multi-client?** v1 supports multiple connected clients but only one
   writer at a time. True multi-writer is out of scope.

## Files

```
docs/future/
├── ai-agents.md                 # This overview
├── phase-1-session-server.md    # TCP server, protocol, security
└── phase-2-agent-kit.md         # MCP tools, CLAUDE.md, demos
```
