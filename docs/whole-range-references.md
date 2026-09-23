# Whole-column and whole-row formula references

Cell formulas and CLI calculations accept `A:A`, `A:B`, `1:1`, and `1:3`.
Endpoints may be absolute (`$A:$B`, `$1:$3`), reversed, or qualified with a
sheet name (`'Sales Data'!A:A`). Copy/fill shifts only relative endpoints on
the named axis. Structural edits adjust that axis and leave the open axis
open. Deleting the entire referenced span produces `#REF!`.

Evaluation bounds the open axis to the referenced sheet's current data extent,
including formulas and spill receivers. It does not scan the grid's unused
trailing rows or columns. Intermediate empty cells remain part of the range;
an empty sheet supplies a single empty row or column to the rectangular range
evaluators. This is a data-bounded range, including for functions that count
blanks or inspect range dimensions.

The stored formula and AST retain the open axis. The dependency graph stores
range subscriptions alongside concrete edges for existing cells, so writes
beyond the previous data extent trigger recalculation. New formula cells gain
concrete edges before ordering; cycle detection includes open ranges even when
the proposed cell is currently empty. Spills participate in the same tracking.

`vgrid sheet inspect --calc`, `vgrid calc`, and hub check calculations use the
engine parser and evaluator. `--headers` supplies a data-row offset for whole
columns; explicit cell ranges and whole-row references retain their coordinates.
String literals such as `"A:A"` are never rewritten.

Regression tests: `crates/engine/tests/whole_range_refs.rs`,
`crates/io/tests/whole_range_roundtrip.rs`, and the CLI inspect tests. The engine
suite includes a deterministic lookup-count test and an opt-in timing comparison
(`cargo test -p visigrid-engine --release --test whole_range_refs -- --ignored --nocapture`).
