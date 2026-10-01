# Tables: engine and structured references

Status: engine foundation and structured-reference implementation, 2026-10-01. This is an engine API and persistence contract, not a released desktop Tables feature. The product plan lives in the Obsidian notes “VisiGrid Tables Spec” and “VisiGrid Tables Research”.

## Model

A `DataTable` identifies an editable rectangle of existing sheet cells. It owns schema, not a second copy of its data. The inclusive rectangle starts with one header row; a header-only table is valid. `TableId` is workbook-unique. `TableColumnId` is scoped to a table. IDs survive renaming, native/JSON saves, and undo/redo. Allocator high-water marks prevent reuse after removal, shrinking, or undo.

Table names share a case-insensitive namespace with named ranges. This slice accepts conservative ASCII names and rejects reference-like names. Column names are case-insensitive, nonempty text; blank/duplicate headers receive deterministic names. Original valid names are reserved before generating suffixes: `Amount, Amount, Amount2` becomes `Amount, Amount3, Amount2`. Names starting with `=` or containing controls are normalized too. Creation previews normalization without mutation.

Table ranges cannot overlap another table, a merged region, pivot output, or an existing array spill. Dynamic arrays cannot spill into a Table, including blank body cells. Ordinary scalar formulas in the body remain ordinary formulas.

## Operations and undo

`Workbook::create_table`, `rename_table`, `rename_table_columns`, `resize_table`, and `remove_table` validate before changing state and return an opaque `TableCommit`. `apply_table_commit(commit, undo)` replays that commit with schema/header and dependent-formula preconditions. A stale commit fails before mutation. No body snapshot is stored; undo retains schema, header preconditions, original typed values of changed headers, and sparse formula-source changes. New dependent formulas that require additional rewrites make replay stale. Creation also captures existing dependent sources so undo restores their originally unbound state.

Creation uses an explicit rectangle whose first row is already the header. It converts normalized headers to text, preserving their explicit formatting. Header insertion and automatic region detection are future UI work.

Resize keeps the top-left corner fixed and changes the bottom/right edges. Surviving columns retain IDs, new columns get fresh IDs, and shrinking leaves released cell values intact. Remove converts dependent structured formulas to absolute A1 references and deletes metadata, preserving values and explicit formats. Conversion refuses referenced empty bodies or unresolved references because they have no lossless A1 representation. The `banded_rows` flag is persisted but is not rendered by this slice.

Headers change through the schema API. Low-level cell setters refuse direct header value writes; tracked workbook setters return no recalculation delta for a refused write. Session batches and operation plans reject header writes during preflight, so they cannot report a successful partial edit. Hosts must use this preflight before other multi-cell editing flows are exposed.

Structural edits entirely before a Table move its bounds with its cells. Edits after it leave its bounds unchanged. Insertions within a Table and deletions intersecting it are temporarily refused. Use explicit resize/remove for schema changes. This keeps the existing structural undo contract intact until table-aware row/column history is integrated. Table-bearing sheet duplication also temporarily refuses instead of duplicating IDs. Sheet removal/restoration updates name reservations and rejects conflicting restoration. Removing a sheet with Table references from other sheets is refused until sheet history can capture those rewrites.

## Structured formulas

The supported subset follows [Microsoft's structured-reference syntax](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables):

| Form | Meaning |
| --- | --- |
| `Sales`, `Sales[#Data]` | All body rows and columns |
| `Sales[Amount]` | One body column |
| `Sales[[Qty]:[Amount]]` | Contiguous body column span |
| `Sales[#Headers]`, `Sales[#All]` | Header row, or headers plus body |
| `Sales[[#Headers],[Amount]]` | Section restricted to a column |
| `[@Qty]`, `[@[Unit Price]]` | Current row, inferred Table context |
| `Sales[@Qty]` | Current row of an explicit Table on the same sheet |
| `[Amount]` | Body column of the Table containing the formula |

Section selectors may also restrict a contiguous column span. Apostrophe escapes `[`, `]`, `#`, `@`, and apostrophe in header names. Names are case-insensitive. Unknown Tables/columns yield `#NAME?`; reversed spans yield `#REF!`; invalid current-row context yields `#VALUE!`. Totals rows, column/section unions, and external-workbook syntax are unsupported.

Examples: `=SUM(Sales[Amount])` works from another sheet; `=[@Qty]*[@Price]` is an ordinary scalar formula in a Table row. Range-consuming functions receive live range descriptors. Standalone nonempty column expressions spill outside Tables; spills into Tables remain prohibited. A header-only Table has zero records: `SUM`/`COUNT` return zero, `ROWS` returns zero, and aggregate errors retain the engine's existing empty-input behavior. A standalone empty Table array produces `#CALC!`.

Stored expressions remain symbolic. Evaluation resolves names against the live schema; dependency extraction lowers references to existing indexed range dependencies rather than expanding per-cell edges. Value edits recalculate incrementally. Schema edits currently rebuild dependencies and recalculate the workbook, including shape-only and empty/nonempty transitions.

Rename rewriting resolves the old schema and follows stable column IDs, including atomic column-name swaps. It changes only source reference spans, preserving strings, whitespace, and grouping elsewhere in the formula. Removing a referenced column writes a permanent `#REF!`; recreating its name does not repair that reference. Undo restores the original source. Local references in released cells become explicitly Table-qualified so they cannot silently adopt another Table's context.

Copy/fill preserves structured tokens and adjusts ordinary A1 references in the same formula. This first subset does not implement Excel's horizontal structured-column fill shifts. Cross-workbook paste provenance, conditional-format/validation formula rewrites, autocomplete, and reference highlighting require integration before those editing flows advertise Table support.

## Persistence

Native `.sheet` workbook saves (ordinary, metadata, full) store a strict versioned `tables` metadata catalog. The single-sheet native API routes table-bearing sheets through the workbook format. Invalid metadata, duplicate identities/names, invalid bounds, ownership collisions, and headers that disagree with schema cause load failure. Missing catalog in an old native file means no tables. Allocators persist even after the last table is removed.

Full JSON emits version 3 only when a workbook/sheet has Table history, with a required `table_catalog`. Ordinary exports retain existing v1/v2 output. Older JSON readers reject v3 rather than silently dropping definitions.

Native semantic fingerprints use v3 for Table-bearing workbooks, including Table names, sheet ownership, bounds, and column names. Presentation and allocator history are excluded. Table-free workbooks retain existing v2 fingerprints.

**Release gate:** old native readers may ignore this new metadata and drop it when saving. A minimum-reader/required-feature compatibility policy is still needed before public Tables authoring ships. XLSX/ODS/CSV and web/cloud editing do not yet preserve table definitions.

## Next slices

1. Desktop creation preview, header editing, table styling/context commands, and history integration; all paste/fill/clear paths must preflight ownership atomically.
2. Table-aware append and structural undo, calculated-column rules and visible exceptions.
3. Table-backed pivot sources with stable field IDs, explicit refresh, and stale-state feedback.
4. XLSX interoperability subset, required-feature file compatibility, web/cloud preservation, and export-loss messaging.
5. Local table views and visible-record paste after the current row-view constraints are addressed.

No new keyboard shortcut or desktop command is exposed by this foundation. Existing PivotTables remain a separate feature.

## Verification

Behavior tests are in `crates/engine/tests/tables.rs`, `crates/engine/tests/structured_tables.rs`, and `crates/io/tests/tables.rs`; session-host and operation-plan tests cover atomic header rejection. Run:

```sh
cargo test -p visigrid-engine -p visigrid-io -p visigrid-session-host
```

The structured-reference regressions cover syntax/escaping, row context, cross-sheet range consumers, incremental value edits, shape changes, rename/column swaps, destructive edits and stale undo, conversion, empty Tables, cycles, copy/fill, structural moves, formula grouping, default Table names, native/JSON reopening, and semantic fingerprints. Desktop UI and XLSX table interoperability are not exercised by this engine slice.

Verified 2026-10-01: 1,279 tests passed, zero failed, 24 existing tests ignored across these packages. This includes 16 structured-reference integration tests and two new IO regressions, in addition to the foundation's 20 Tables tests.
