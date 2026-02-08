# Array Semantics (Phase 2)

Status: **Deferred** — deliberately not implemented in Phase 1.

## Context

Excel supports implicit array arithmetic in expressions like:

```
=SUM(A1:A10 * B1:B10)
=SUM(N67:N71 * ($E7:$E15 / (1+$E7:$E15)))
```

These work because Excel:
1. Evaluates range expressions as arrays
2. Applies elementwise arithmetic (`*`, `/`, `+`, `-`)
3. Creates intermediate arrays
4. Reduces with aggregate functions (`SUM`, `AVERAGE`, etc.)

VisiGrid currently rejects these with `#VALUE! Array arithmetic not supported`. This is intentional — determinism and auditability come first.

## Current State

- **SUMPRODUCT** is implemented as a targeted escape hatch for the common `SUM(range*range)` pattern
- Remaining array arithmetic errors in real-world imports (e.g. SaaS financial models) are ~168 out of 6,334 formulas
- Error message is short and honest, no nested wrapping

## What Full Array Support Requires

### Type system

- `Value::Array(Array2D)` — 2D typed array as a first-class value
- Arrays flow through the expression evaluator, not just inside specific functions

### Elementwise operations

Binary ops on arrays:
```
Array(3x1) * Array(3x1) → Array(3x1)  // elementwise
Array(3x1) * Scalar     → Array(3x1)  // broadcast scalar
Array(3x1) * Array(2x1) → #VALUE!     // shape mismatch
```

### Broadcasting rules

Define explicitly and document:
- Scalar broadcasts to any shape
- 1xN broadcasts to MxN (and vice versa)
- Mismatched shapes → error (no NumPy-style broadcasting)

### Shape validation

All intermediate arrays must have validated shapes before operations proceed.

### Deterministic reduction

Aggregate functions (`SUM`, `AVERAGE`, `MIN`, `MAX`, etc.) must accept `Value::Array` and reduce deterministically in row-major order.

### Spill-aware evaluation

If a formula produces an array and is not inside an aggregate:
- Spill to adjacent cells (Excel 365 dynamic arrays)
- Or error if spill region is occupied
- VisiGrid already has spill infrastructure (`spill_info`, `spill_parent`)

## Design Principles

1. **Deterministic** — same inputs always produce same outputs, same order
2. **Explicit** — no silent coercion or locale-dependent behavior
3. **Auditable** — intermediate values can be inspected
4. **Conservative** — reject ambiguous cases rather than guess

## What NOT to Do

- Don't silently coerce shapes
- Don't auto-wrap in SUMPRODUCT
- Don't pick first element from arrays
- Don't flatten arrays
- Don't depend on iteration order of sparse structures

## Migration Path

1. **Phase 1 (done):** SUMPRODUCT as explicit escape hatch, clear error messages
2. **Phase 2:** `Value::Array` type, elementwise ops, shape validation
3. **Phase 3:** Spill semantics, dynamic array formulas, `SEQUENCE`/`SORT`/`UNIQUE` producing arrays

## Related

- `execution-contracts-spec.md` — deterministic evaluation guarantees
- `dependency-graph-spec.md` — recalculation order
- SUMPRODUCT implementation in `crates/engine/src/formula/eval_math.rs`
