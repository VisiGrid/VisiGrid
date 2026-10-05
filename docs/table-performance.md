# Table edit performance

## Reproducing the measurement

From the app workspace:

```sh
cargo run --release -p visigrid-engine --example table_edit_bench -- 10000 100000
```

The example defaults to those two row counts and three samples per case. Set
`TABLE_BENCH_SAMPLES` to change the sample count; pass other row counts as
arguments. It uses the repository's release profile, including size optimization
and LTO. Setup is excluded from the component timings. Each repetition starts
from the same immutable workbook. Median, minimum and maximum are reported.

Each fixture contains a three-column Table (`Region`, `Amount`, `Result`), a
descending Amount sort, a West-only checkbox filter, native totals and a
cross-sheet `SUM` over Result. One case stores Result values; the other has one
formula per body row. Worksheets retain their full 1,048,576-row dimensions,
including blank rows represented by the view. The edit changes a visible
record's Amount from 2 to 12345 and moves its sorted position.

The benchmark checks canonical visibility, edited values, dependent formulas,
totals, one authored-cell patch, undo/redo and original-workbook nonmutation.
It times these engine components individually:

- Build the sorted/filtered view.
- Clone the candidate workbook.
- Apply a tracked write in a batch and recalculate dependents.
- Rebuild every saved view to validate the candidate.
- Capture a guarded batch commit and prepare its undo/redo candidates.
- Publish a snapshot while retaining monotonic revisions.

**This is not an end-to-end desktop latency measurement.** Ordinary typing uses
`TableCellsCommit`, a target-only history path, rather than
`capture_guarded_batch`. The latter covers automation batches and structural or
metadata changes such as manual row visibility. Desktop target/layout checks,
additional view builds, focus handling and rendering are outside this example.
Do not sum these components and present that as typing latency.

RSS is process resident memory, not live heap allocation. Peak RSS is cumulative
across cases and includes setup and simultaneous replay candidates. It is zero
on platforms unsupported by the existing RSS helper. Allocator trimming before
each baseline is supported only on glibc. Run cases in separate processes for
independent peak-memory measurements.

## Local baseline, 2026-10-04

Linux x86-64, Intel Core i7-1270P, 32 GiB RAM, Rust 1.98.1, three samples per
case. Other desktop workloads were running, so timings are diagnostic rather
than a release performance guarantee. Baseline engine: `88e142f`.

| Component, median milliseconds | 10k values | 10k formulas | 100k values | 100k formulas |
| --- | ---: | ---: | ---: | ---: |
| Build view | 27.02 | 21.37 | 56.63 | 56.79 |
| Clone workbook | 0.01 | 0.01 | 0.02 | 0.02 |
| Write and incremental recalc | 0.02 | 5.00 | 0.02 | 59.10 |
| Validate saved views | 25.86 | 23.87 | 60.78 | 50.56 |
| Capture guarded batch | 988.19 | 755.07 | 9277.36 | 6903.71 |
| Prepare guarded undo | 377.72 | 260.96 | 2827.86 | 2589.41 |
| Prepare guarded redo | 384.55 | 261.18 | 4351.10 | 2558.45 |
| Publish snapshot | 0.01 | 0.02 | 0.03 | 0.05 |

The process reached a cumulative peak of 1877.5 MiB. Snapshot cloning was cheap
with the current shared storage. Guarded history scanning and fingerprinting
were the dominant measured costs. The old fingerprint implementation retained a
JSON tree for every stored cell in a sorted map before hashing it.

## Optimization scope

Fingerprinting now sorts cell coordinates, then serializes and hashes one cell
at a time. It retains the exact prior encoding: position, encoded length and
authored-cell signature. A compatibility regression compares the old and new
encoders across value types, insertion orders, formulas, formatting, comments,
styles, frozen formulas and derived spill state. It also checks that authored
changes, including signed zero and cell removal, still change the fingerprint.

Candidate validation, full-workbook stale checks and sparse history contents
remain unchanged. This removes the retained JSON trees; it does not eliminate
full scans or the cost of comparing cell signatures. Repeated view construction
and further history optimizations need separate measurements and correctness
checks.

## After streaming fingerprints

Same machine, release profile, fixture sizes and three-sample command:

| Component, median milliseconds | 10k values | 10k formulas | 100k values | 100k formulas |
| --- | ---: | ---: | ---: | ---: |
| Build view | 37.92 | 39.49 | 78.94 | 56.45 |
| Clone workbook | 0.01 | 0.01 | 0.02 | 0.02 |
| Write and incremental recalc | 0.02 | 5.65 | 0.02 | 67.36 |
| Validate saved views | 37.98 | 37.54 | 76.95 | 55.85 |
| Capture guarded batch | 972.58 | 866.42 | 8016.75 | 5443.55 |
| Prepare guarded undo | 391.58 | 346.64 | 2817.92 | 2528.07 |
| Prepare guarded redo | 328.94 | 339.32 | 2709.11 | 2816.26 |
| Publish snapshot | 0.01 | 0.02 | 0.02 | 0.05 |

All benchmark correctness assertions passed. Cumulative process peak RSS fell
from **1877.5 MiB to 209.6 MiB**, about 89%. The value-only 100k case did not
exceed the 64.0 MiB peak already reached by the preceding 10k formula case.
The 100k formula fixture's baseline RSS stayed nearly identical (83.7 versus
83.6 MiB), consistent with removing temporary fingerprint allocations rather
than changing workbook contents or representation.

Timing ranges overlap and unmodified view-building timings vary substantially,
so this run supports a memory improvement, not a general latency guarantee.
Large guarded captures remain multi-second operations. Future work should
profile signature comparison/serialization and repeated full scans, retain the
same stale-history coverage, and separately measure actual desktop typing,
paste, undo and rendering on representative workbooks.

Validation: 1,883 engine and desktop tests passed, with 18 existing ignores.
The native desktop build and diff checks passed. No live UI, Microsoft Excel
or cross-platform performance checks were performed in this measurement pass.

## Large Table conversion and compact history

Run each fixture in a separate process, after building the example:

```sh
cargo build --release -p visigrid-engine --example table_conversion_bench
target/release/examples/table_conversion_bench 10000
target/release/examples/table_conversion_bench 100000
target/release/examples/table_conversion_bench 1000000
```

The fixture has ten columns: eight numeric and two calculated, native totals,
and a cross-sheet structured `SUM`. Worksheets use their full dimensions. Setup
is excluded. Conversion, undo and redo are timed separately. Assertions check
every converted body formula's result, totals, the dependent summary, replay,
subsequent recalculation and atomic stale-history refusal. History-only RSS is
measured after dropping the workbook and trimming the allocator, then again
after dropping history. That difference is diagnostic process memory, not an
allocation counter. This benchmark still excludes desktop dispatch/rendering.

The pre-optimization baseline (`6d15354`) scans the complete rewrite list for
each formula during Table-operation validation. That quadratic check now uses
an indexed set of cell identities. Ordinary source checks and new-dependent
refusal remain in place. Guarded capture compares borrowed authored fields
instead of cloning cells and constructing JSON trees for unchanged values.
Fingerprint serialization reuses a buffer and a bounded format cache while
retaining the exact original encoding. Formula changes with unchanged
presentation retain only their before/after sources; other changes retain
complete authored images. The full-workbook fingerprint still protects all
formatting, comments, styles, frozen sources and other authored state.

Conversion no longer inherits the ordinary 100,000-cell batch limit and no
longer retains a duplicate list of formula/rule rewrites alongside guarded
history. Other guarded operations retain their existing limits. Replay reparses
source-only formula patches and rebuilds calculation caches and spills on its
candidate before publishing. Full scans,
recalculation and view validation remain; this is not an off-thread execution
or cancellation implementation.

### Conversion measurements, 2026-10-05

Same Linux x86-64 Intel Core i7-1270P machine with 32 GiB RAM and Rust 1.98.1,
using the repository release profile. One sample per size and implementation,
each in a separate process. Other workloads were present; these are diagnostic
measurements, not latency guarantees. In particular, setup times differed
substantially between runs. Baseline setup and conversion used the same fixture
and the pre-optimization engine executable retained before these changes.

| Body rows | Baseline conversion, seconds | New conversion, seconds | New undo, seconds | New redo, seconds |
| --- | ---: | ---: | ---: | ---: |
| 10,000 | 1.718 | 1.173 | 0.533 | 0.518 |
| 100,000 | 66.378, then refused | 12.103 | 5.883 | 6.401 |
| 1,000,000 | 7,581.916, then refused | 135.101 | 69.981 | 61.485 |

Both baseline refusals were the 100,000-changed-cell history limit, after
performing the expensive candidate conversion. The baseline verified that its
refusal left the original Table and formulas intact. All three new runs passed
the full success-path assertions, including every body result, totals,
cross-sheet results, undo/redo, a subsequent value edit and stale-undo refusal.

| Body rows | Baseline peak through conversion, MiB | New peak through conversion, MiB | Baseline retained history, approximate MiB | New retained history, approximate MiB |
| --- | ---: | ---: | ---: | ---: |
| 10,000 | 63.9 | 50.5 | 19.6 | 4.1 |
| 100,000 | 475.5 | 440.8 | unavailable: refused | 32.9 |
| 1,000,000 | 4,738.0 | 4,736.5 | unavailable: refused | 290.3 |

Retained-history estimates subtract RSS after dropping history from RSS with
only history retained, after allocator trimming in each case. The 10k sample
fell about 79%. The million-row workbook's baseline RSS was about 1,923 MiB in
both runs. Peak memory at that size did not improve materially: temporary
workbook/calculation work still dominates. Including undo/redo, the new process
peaked at 4,864.9 MiB. The result removes a quadratic validation scan and the
conversion history ceiling; it does not establish interactive performance or
solve the outstanding desktop scheduling/cancellation work.

Validation: the combined engine, I/O, desktop, session-host, browser-adapter
and CLI run passed 3,501 tests across 99 suites, with zero failures and 35
existing ignores. New tests compare borrowed cell equality and fingerprint
encoding with the original JSON semantics, exercise the bounded format cache,
preserve malformed formula variants and presentation, replay mixed full-cell
and source-only patches with changing spills, and convert/undo/redo 100,002
formula rewrites while retaining the ordinary batch limit and stale-comment
refusal. Existing Table/schema, conversion, persistence and desktop History
rewind regressions passed. The launchable native desktop build and diff checks
passed. Live UI, real Excel and cross-platform verification remain outstanding.
