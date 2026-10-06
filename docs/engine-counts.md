# Engine instruction counts

A deterministic performance gate for the engine's hot paths. Phase 2 of the
web performance plan.

## Why instructions, not milliseconds

Wall-clock time is what users feel, and it is the field metric the web app
reports to PostHog. It is also noisy: a CI runner shares its CPU, caches
warm and cool, and a 3% change is lost in the spread. A millisecond figure
cannot fail a pull request without failing honest ones.

Instructions executed, counted by Valgrind's callgrind, do not depend on
load. The same binary on the same input gives the same count to well under
a percent. One run is a measurement, so a count can have a ceiling and the
ceiling can be enforced.

Instructions are a proxy. Before trusting a scenario, prove once that
driving its count down also drives its wall-clock time down, and record both
numbers in the plan's Ratchets table. A scenario that fails that proof
should be removed, not kept as a hill to climb.

## Scenarios

`crates/io/examples/engine_counts.rs`, one `measured_*` function each. Only
the instructions inside that function are counted; building the input is
free.

| Scenario | What is counted | Why it matters |
| --- | --- | --- |
| `parse_formulas` | Writing 1,000 formula cells, which parses each | Open and paste paths |
| `full_recalc` | Dependency graph and full recompute of 20,000 formulas and one SUM over the column | Open a workbook; server recalculation |
| `incremental_edit` | One edit to a cell with 1,000 direct dependents | The keystroke path |
| `json_roundtrip` | Export 5,000 rows to canonical visigrid-json and read it back | Save and open on every client |

## Running

Linux with valgrind:

```sh
scripts/engine-counts.sh target/counts
scripts/engine-counts-gate.py target/counts/results.json
```

Without valgrind the scenarios still run as a plain binary and check their
results, which is enough to know a scenario is correct:

```sh
cargo run --locked -p visigrid-io --no-default-features --example engine_counts -- full_recalc
```

## The ratchet

Ceilings are in `benches/engine-counts.ceilings.<arch>.json`, one per
scenario, per CPU architecture (counts differ between x86_64 and arm64).

- A pull request that touches `crates/engine` or `crates/io` runs the counts
  and fails if any scenario exceeds its ceiling by more than 1%.
- The nightly run on main lowers any ceiling the count came in under and
  opens a PR with the new file. Ceilings only go down.
- A new scenario has no ceiling; the gate reports it and passes, and the
  next nightly seeds it.

If a change legitimately needs more instructions, raise the ceiling in the
same PR and say why in the commit message. That is the only way up.
