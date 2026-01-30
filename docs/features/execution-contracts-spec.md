# Execution Contracts Specification

> **Status:** Mostly implemented (on `feature/analyze-ai`)
> **Release:** 0.4.1
> **Prerequisite:** Analyze with AI (same branch)
> **Nature:** Language + UX release, not architecture

---

## The Point

Every AI action in VisiGrid operates under a named contract. The contract tells the user — before, during, and after — exactly what AI can and cannot do.

This is not a technical formalism. There are no contract structs, no hashing, no enforcement engine. This is a **naming release**: the moment VisiGrid stops being "an app with AI features" and becomes "an app where AI operates under rules."

What ships:
- Execution Contract badges everywhere AI appears
- Contract name and write scope in every transparency panel
- Internal constants that make the boundary grep-able and testable

What does not ship:
- No `ExecutionContract` struct
- No cryptographic hashing
- No proposal engine
- No formal verification

---

## Contracts Defined

| Contract | Identifier | Write scope | Used by |
|----------|-----------|-------------|---------|
| Read-only v1 | `read_only_v1` | None | Analyze, Diff Summary, Diff Entry Explain |
| Single-cell write v1 | `single_cell_write_v1` | Active cell | Insert Formula |

These are the only two contracts that exist in 0.4.1. A third (`proposed_changes_v1`) will arrive with Propose Changes. Do not pre-define it.

---

## Current State (what's already done)

The `feature/analyze-ai` branch already ships most of this:

| Touchpoint | Contract badge | Sent-to-AI panel | Status |
|------------|---------------|------------------|--------|
| Insert Formula dialog header | "Single-cell write" | Contract + Write scope rows | Done |
| Analyze dialog header | "Read-only" | Contract + Write scope rows | Done |
| Diff Summary (Generate Summary button) | — | — | **Gap** |
| Diff Entry Explain (per-entry explain) | — | — | **Gap** |
| Copy Details (diagnostic dump) | — | — | **Gap** |

Constants defined in `gpui-app/src/ai/client.rs`:
```rust
pub const ANALYZE_CONTRACT: &str = "read_only_v1";
pub const INSERT_FORMULA_CONTRACT: &str = "single_cell_write_v1";
```

---

## What Needs to Ship (the gaps)

### 1. Diff Summary — contract badge

The "Generate Summary" button in the Explain Differences dialog (`inspector_panel.rs:4321`) calls AI to summarize all changes between two snapshots. This is read-only — it writes nothing.

**Change:** Add a muted contract line below the "AI Summary" header:

```
AI Summary
Read-only ⓘ   Summary only — no cells modified.
```

The ⓘ icon shows the one-line definition on hover. Styled identically to the Analyze dialog badge: muted text, subtle background, not dismissible.

**Location:** `inspector_panel.rs`, inside the `ai_summary_available` conditional block, after the "AI Summary" label (line 4343).

### 2. Diff Entry Explain — contract badge

Individual entry explanations (the "Explain" button next to each diff entry) also call AI. Same contract.

**Change:** When an explanation is displayed below an entry, prefix it with a contract indicator. Keep it minimal — this is inline, not a dialog:

```
Read-only — The AI read your data to generate this explanation.
[explanation text here]
```

Or, if space is tight, a single muted line above the explanation text:

```
Read-only ⓘ
```

Either form uses the standard "Read-only" label. The ⓘ tooltip provides the one-line definition.

**Location:** `inspector_panel.rs`, wherever entry explanations are rendered (inside `render_diff_section`).

### 3. Copy Details — include contract in diagnostic dump

The `ask_ai_copy_details()` function (`dialogs.rs:1805`) copies diagnostic info to clipboard for bug reports. It already includes the verb label but not the contract identifier.

**Change:** Add a contract line to the diagnostic output. Include both the grep-able identifier and the human label so the dump is self-documenting for both developers and users:

```
=== Insert Formula AI Diagnostic Details ===

Contract: single_cell_write_v1 (Single-cell write)
Write scope: Active cell
Provider: openai
Model: gpt-4o-mini
...
```

For read-only verbs:

```
=== Analyze AI Diagnostic Details ===

Contract: read_only_v1 (Read-only)
Write scope: None
...
```

**Location:** `dialogs.rs`, `ask_ai_copy_details()` method, after the verb label line.

### 4. Diff Summary / Entry Explain — add contract constant

The diff features currently use AI but don't reference a named contract. Add a constant:

```rust
pub const DIFF_EXPLAIN_CONTRACT: &str = "read_only_v1";
```

This is the same value as `ANALYZE_CONTRACT` — that's correct. They share the same contract. The constant exists so it's grep-able and so the diff code explicitly declares what contract it operates under.

**Location:** `gpui-app/src/ai/client.rs`, alongside existing constants.

---

## What Does Not Change

- **AI Settings dialog** — This is configuration, not execution. No contract badge needed.
- **AI doctor / diagnostics** — Shows capabilities, not contracts. Capabilities answer "what can this provider do?" Contracts answer "what will this action do?" Different questions.
- **Status bar** — No passive AI indicators exist. Don't add them.
- **System prompts** — Already correct. No changes needed.

---

## Design Rules

1. **Badge is always visible.** Never collapsed, never scrolled away, never behind a toggle. If AI is running, the contract is showing.

2. **Badge language is human, not technical.** Users see "Read-only" and "No cells will be modified." They do not see `read_only_v1`. The identifier lives in the Sent-to-AI panel and diagnostic dumps.

3. **Don't over-badge.** The diff features are smaller touchpoints — a single muted line is enough. Don't add dialog-scale badges to inline explanations.

4. **Constants are cheap insurance.** Even when two features share the same contract value, give each its own constant. This costs nothing and makes auditing trivial:
   ```
   grep -r "CONTRACT" gpui-app/src/ai/
   ```

5. **No contract struct yet.** The temptation will be to create `struct ExecutionContract { name, write_scope, ... }`. Don't. Constants and badges are sufficient for 0.4.1. Structs arrive when Propose Changes needs them — not before.

6. **One-line definition available on hover.** The first time a user sees "Execution Contract" they should be able to learn what it means without leaving the dialog. Add a small ⓘ icon or tooltip next to the contract badge with the text: *"Execution Contracts define what AI is allowed to read and write."* This is not a tutorial — it's a single sentence for orientation. Present on all contract badges, triggered on hover.

7. **Consistent phrasing.** The human-facing label for the read-only contract is always **"Read-only"** — never "Read only" (no hyphen), never "Read-Only" (capital O), never just "Read". Grep the UI strings before shipping and normalize.

---

## Tests

### Unit tests

1. `ANALYZE_CONTRACT` equals `"read_only_v1"` — **already passing**
2. `INSERT_FORMULA_CONTRACT` equals `"single_cell_write_v1"` — **already passing**
3. `DIFF_EXPLAIN_CONTRACT` equals `"read_only_v1"` — **new**
4. All contract constants are non-empty strings

### Visual verification

5. Insert Formula dialog shows "Single-cell write" badge in header
6. Analyze dialog shows "Read-only" badge in header
7. Diff Summary section shows read-only contract line
8. Diff Entry explanation shows contract indicator
9. Copy Details output includes contract identifier with human label and write scope
10. Sent-to-AI panel shows Contract and Write scope rows for both verbs
11. ⓘ icon on contract badges shows tooltip: "Execution Contracts define what AI is allowed to read and write."

---

## Implementation Scope

This is a small delta on top of the `feature/analyze-ai` branch:

| File | Change | Size |
|------|--------|------|
| `gpui-app/src/ai/client.rs` | Add `DIFF_EXPLAIN_CONTRACT` constant | 1 line |
| `gpui-app/src/views/inspector_panel.rs` | Badge below "AI Summary" header | ~8 lines |
| `gpui-app/src/views/inspector_panel.rs` | Contract line above entry explanations | ~5 lines |
| `gpui-app/src/dialogs.rs` | Contract + write scope in copy details | ~4 lines |
| `gpui-app/src/ai/mod.rs` | Export new constant | 1 line |

Total: ~20 lines of meaningful code. The rest is already done.

---

## What This Release Says

After 0.4.1, every AI interaction in VisiGrid has:

- A visible contract name
- An explicit write scope
- A transparency panel that proves both

The user's mental model becomes:

> "AI in VisiGrid tells me what it's allowed to do before it does it."

That's the category. That's what Propose Changes inherits. That's what makes the next release possible without trust regression.

---

## Success Metric

> Open every AI feature in VisiGrid. Can you see, without clicking anything, what AI is allowed to do and what it is not?

If yes for all features — ship it.
