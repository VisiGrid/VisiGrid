# Future Features: Systems of Record + AI Reconciliation

This document captures feature ideas for VisiGrid's evolution from spreadsheet to **local-first computation surface**.

---

## Core Positioning

VisiGrid is not building:
- AI models
- Finance templates (amortization, schedules, etc.)
- Domain-specific calculators

VisiGrid is building:
- A local-first computation surface that AI + APIs can act through
- A programmable decision engine people trust

This distinction determines whether VisiGrid becomes another spreadsheet clone or something meaningfully different.

---

## Feature A: Systems of Record Connectors

### What It Is

Authoritative data ingress from trusted systems. Not "API integrations" marketing fluff.

Examples of Systems of Record:
- ERP (NetSuite, SAP, QuickBooks)
- Stripe / payment processors
- Internal ledger / database exports
- Bank payment feeds
- Payroll systems

### Design Principles (Hard Rules)

A system-of-record connector must be:

| Property | Rationale |
|----------|-----------|
| **Read-only** | Never write back to source systems |
| **Credential-scoped** | Auth stored locally, not in cloud |
| **Locally stored** | Data lives on user's machine |
| **Explicitly refreshed** | No silent syncs, user controls when |
| **Snapshot-preserving** | Every pull is versioned with timestamp |
| **Schema-stable** | Predictable table structure |

### Mental Model

Treat every SoR as: **A signed statement of facts at time T**

Not a live feed. Not a sync engine. This is how auditors think.

### Example

```
Source: Stripe
Scope: Charges + Payouts
Auth: Stored locally (encrypted)
Refresh: Manual
Snapshot: 2026-01-17T10:14
Output: Table
```

That snapshot is sacred.

### What We Deliberately Do NOT Do

- No auto-write-back
- No silent refresh
- No implicit joins
- No hidden transformations

If Stripe says $10,000, that's what the table says - even if it's wrong.

### UX Consideration

Start with a dedicated **Sources panel** (sidebar) rather than formula syntax like `=API("stripe_payments")`. The formula syntax is elegant but may be premature. Validate the mental model first, syntax can come later.

---

## Feature B: AI as Reconciliation Intelligence

### What It Is

AI answers: **"Given these sheets, what should I look at?"**

Not: "What should I change?"

### The Contract (Non-Negotiable)

**AI can:**
- Propose operations
- Label results
- Explain diffs
- Flag anomalies
- Summarize discrepancies

**AI cannot:**
- Mutate cells without confirmation
- Invent values
- Hide steps
- Fetch data directly
- Refresh APIs
- Edit sheets

AI produces **plans**, not actions. Execution is a separate, user-approved step.

### Allowed Operation Types

| Operation | Description |
|-----------|-------------|
| Compare | Match records across sources |
| Reconcile | Identify what lines up and what doesn't |
| Explain | Describe why differences exist |
| Flag | Mark anomalies or violations |
| Diff | Show changes between snapshots |
| Summarize | Aggregate findings |

### Forbidden Operations

- "Build me an amortization table"
- "Fix this sheet"
- "Auto-fill these cells"

These destroy provenance and trust.

### Example Workflow

User prompt:
> "Compare Stripe payouts to ERP cash receipts and flag timing differences vs true mismatches."

AI output:
- Proposed join keys
- Tolerance windows (e.g., dates within 3 days)
- Mismatch categories
- Explanation of each discrepancy

User clicks **Run**.

VisiGrid executes deterministically.

Result:
- New diff sheet
- Annotations on flagged rows
- Explanation panel

**AI never touched the data. API never reasoned. User stayed in control.**

---

## Why This Separation Matters

### The Copilot Problem

Reddit sentiment on Excel + Copilot is deeply unimpressed:

1. **"It hallucinates formulas"** - invents functions that don't exist, confidently gives almost-right answers
2. **"Slower than just doing it myself"** - prompting + waiting + correcting + validating = negative ROI
3. **"I don't know what it changed"** - no diff, no execution trace, no semantic rollback
4. **"Useless for real business mess"** - works only on clean demo sheets
5. **"Security/compliance veto"** - data leaving machine, unclear retention

### Why Copilot Is Structurally Broken

| Sacred Spreadsheet Rule | Copilot Violation |
|------------------------|-------------------|
| Spreadsheets are systems of record (determinism, traceability) | Copilot is probabilistic |
| Power users think procedurally (exact operations) | Copilot thinks "close enough" |
| Trust > cleverness (being wrong once is worse than being slow) | Copilot optimizes for delight |

### The Opening for VisiGrid

Reddit is not anti-AI. They are anti-**opaque** AI.

What they want:
- AI that proposes, not edits
- AI that explains, not guesses
- AI that flags, not fixes
- AI that never touches data silently

---

## How They Compose

The magic happens only when the user explicitly connects them.

```
[SoR: Stripe]          [Local: check_log.csv]
      │                        │
      ▼                        ▼
   Sheet A                  Sheet B
      │                        │
      └──────────┬─────────────┘
                 │
                 ▼
         User Prompt:
    "Compare and flag mismatches"
                 │
                 ▼
           AI Response:
         (comparison plan)
                 │
                 ▼
          User: [Run]
                 │
                 ▼
      Deterministic Execution
                 │
                 ▼
         Diff Sheet + Annotations
```

---

## Site Positioning

### Systems of Record

> **Import authoritative data from ERPs, payment processors, and ledgers - without sync risk.**
>
> Connect once. Refresh when you decide. Every snapshot preserved.

### AI Reconciliation

> **AI helps you reason about your data - it never becomes your data.**
>
> Compare systems, explain differences, approve actions. Every step visible.

Or more punchy:

> **AI, without the magic tricks**
>
> Use AI to compare, reconcile, and explain data - not to rewrite it. Every result is auditable, local, and reversible.

---

## What NOT to Build (Discipline)

- Prebuilt finance templates
- Domain models
- "AI formulas"
- SaaS-style connector explosion
- Collaboration (yet)

These are all downstream. The power is **neutrality**.

---

## Execution Roadmap

### Phase 1: Systems of Record
- [ ] SoR source definition schema
- [ ] Local encrypted credential storage
- [ ] Manual refresh mechanism
- [ ] Table materialization
- [ ] Snapshot history with timestamps

### Phase 2: AI Reasoning
- [ ] Structured operation schema (Compare, Flag, Explain, etc.)
- [ ] Read-only data access
- [ ] Diff generation
- [ ] Explanation panel UI

### Phase 3: Composition
- [ ] "Run plan" approval UX
- [ ] Audit log
- [ ] Operation replay

### Golden Path (Ship First)
**Payment reconciliation: Stripe vs CSV ledger**

Everyone understands this. If it feels amazing, we win.

---

## The Bigger Vision

If done right, VisiGrid is not "Excel but faster."

It's: **A local, programmable truth engine for messy business data.**

That's a big idea. And it's coherent.
