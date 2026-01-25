# VisiGrid Product Doctrine

## Category

**VisiGrid is an Explainable Spreadsheet.**

Users can:
- Trust that values are correct
- Understand *why* values are what they are
- Predict what changes will affect
- Defend their work to others

## PR Kill-Question

Before merging any PR, ask:

> "Does this make causality, impact, or recomputation more visible?"

If the answer is no, the PR needs justification or rejection.

## Non-Goals (Hard Cuts)

These will never be in VisiGrid:

1. **Implicit execution** — No auto-running macros, no hidden side effects
2. **Second computation model** — One formula engine, one evaluation order
3. **Invisible state** — If it affects output, it must be inspectable
4. **Magic defaults** — Prefer explicit over convenient

## Tiering Principle

| Question | Tier |
|----------|------|
| Is this value correct? | Free |
| Why is this value correct? | Paid |
| Can I prove it to someone else? | Paid |

**Rule:** Trust is free. Explanation is paid.

**Clarification:** Depth as a scalar is a summary metric (free). Graph structure is explanation (paid). This distinction prevents tier boundary debates.

## PR Checklist

- [ ] Improves explainability (or is infrastructure for it)
- [ ] Does not introduce implicit execution
- [ ] Does not add a second computation model
- [ ] Surfaces state that affects output (no hidden state)

## Known Limitations (Documented, Not Hidden)

- **Dynamic refs** (`INDIRECT`, `OFFSET`) have unknown dependencies at parse time. Graph edges may be incomplete. Impact is marked as "unbounded" for these cells.
- **Cycles** are detected at edit-time, not evaluation-time. Error on entry, not silent `#VALUE` later.

---

*This document is the source of truth for product decisions. Update it when the category evolves.*
