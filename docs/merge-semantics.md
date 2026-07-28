# Merged Cell Semantics

Behavioral contract for merged regions. Code must obey this spec.

## Model

A merged region has one **origin** (top-left cell) and zero or more
**hidden cells**. The origin holds the region's value and format. Hidden
cells are physically empty in storage — merging does not copy or move
values into them, and writes addressed to a hidden cell redirect to the
origin (`merge_origin_coord`).

## Reads — Excel parity (ruled 2026-07-28)

Formula reads of a **hidden cell resolve to EMPTY**, never to the origin:

- `=B1` with A1:B1 merged and A1 = "Hello" → empty (displays as "")
- `=A1+B1+C1` with A1:C1 merged and A1 = 10 → 10 (B1 and C1 contribute 0)
- Cross-sheet refs behave identically (`=Sheet1!B1` → empty)
- Range functions were already consistent with this (`=SUM(A1:C1)` → 10,
  because hidden cells are empty in storage)

History: before 2026-07-28 single-cell refs *redirected* to the origin
(`=B1` → "Hello") while ranges read hidden cells as empty — an internal
inconsistency and a divergence from Excel. The redirect was removed from
formula evaluation (`eval.rs`); the ruling is Excel parity throughout.
Note this changed recompute results for sheets that relied on the
redirect (shipped in the first release after v0.14.0).

## Writes

Writes to a hidden cell redirect to the origin. Unmerging leaves the
value in the origin; formerly hidden cells stay empty.

## Structural operations

Row/column insert/delete shift merge regions with grid-line semantics
(regions grow/shrink when lines are inserted/removed strictly inside
them). Conditional-formatting rule ranges follow the same rules.
