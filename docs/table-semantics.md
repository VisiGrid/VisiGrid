# Tables: engine, structured references, and desktop authoring

Status: desktop authoring implementation, 2026-10-01. This is the staged implementation contract, not a declaration that the complete Tables release is ready. The product plan lives in the Obsidian notes “VisiGrid Tables Spec” and “VisiGrid Tables Research”.

## Model

A `DataTable` identifies an editable rectangle of existing sheet cells. It owns schema, not a second copy of its data. The inclusive rectangle starts with one header row; a header-only table is valid. `TableId` is workbook-unique. `TableColumnId` is scoped to a table. IDs survive renaming, native/JSON saves, and undo/redo. Allocator high-water marks prevent reuse after removal, shrinking, or undo.

Table names share a case-insensitive namespace with named ranges. This slice accepts conservative ASCII names and rejects reference-like names. Column names are case-insensitive, nonempty text; blank/duplicate headers receive deterministic names. Original valid names are reserved before generating suffixes: `Amount, Amount, Amount2` becomes `Amount, Amount3, Amount2`. Names starting with `=` or containing controls are normalized too. Creation previews normalization without mutation.

Table ranges cannot overlap another table, a merged region, pivot output, or an existing array spill. Dynamic arrays cannot spill into a Table, including blank body cells. Ordinary scalar formulas in the body remain ordinary formulas.

## Operations and undo

`Workbook::create_table`, `rename_table`, `rename_table_columns`, `resize_table`, and `remove_table` validate before changing state and return an opaque `TableCommit`. `apply_table_commit(commit, undo)` replays that commit with schema/header and dependent-formula preconditions. A stale commit fails before mutation. No body snapshot is stored; undo retains schema, header preconditions, original typed values of changed headers, and sparse formula-source changes. New dependent formulas that require additional rewrites make replay stale. Creation also captures existing dependent sources so undo restores their originally unbound state.

Creation uses an explicit rectangle whose first row is already the header. It converts normalized headers to text, preserving their explicit formatting. The desktop suggests the current region for a single-cell selection; inserting a header row remains future work.

Resize keeps the top-left corner fixed and changes the bottom/right edges. Surviving columns retain IDs, new columns get fresh IDs, and shrinking leaves released cell values intact. Remove converts dependent structured formulas to absolute A1 references and deletes metadata, preserving values and explicit formats. Conversion refuses referenced empty bodies or unresolved references because they have no lossless A1 representation. `set_table_style` changes persisted banding through the same guarded commit API.

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

Copy/fill preserves structured tokens and adjusts ordinary A1 references in the same formula. This first subset does not implement Excel's horizontal structured-column fill shifts. The editor treats structured selectors as opaque reference tokens, suppressing unrelated function completion and false bracket diagnostics. Column autocomplete, range highlighting, cross-workbook paste provenance, and conditional-format/validation formula rewrites still require integration before those editing flows advertise Table support.

## Persistence

Native `.sheet` workbook saves (ordinary, metadata, full) store a strict versioned `tables` metadata catalog. The single-sheet native API routes table-bearing sheets through the workbook format. Invalid metadata, duplicate identities/names, invalid bounds, ownership collisions, and headers that disagree with schema cause load failure. Missing catalog in an old native file means no tables. Allocators persist even after the last table is removed.

Full JSON emits version 3 only when a workbook/sheet has Table history, with a required `table_catalog`. Ordinary exports retain existing v1/v2 output. Older JSON readers reject v3 rather than silently dropping definitions.

Native semantic fingerprints use v3 for Table-bearing workbooks, including Table names, sheet ownership, bounds, and column names. Presentation and allocator history are excluded. Table-free workbooks retain existing v2 fingerprints.

**Release gate:** old native readers may ignore this new metadata and drop it when saving. A minimum-reader/required-feature compatibility policy is still needed before public Tables authoring ships. XLSX/ODS/CSV and web/cloud editing do not yet preserve table definitions.

## Desktop authoring

- **Insert → Table**, **Create Table** in the command palette, and platform-primary **T** open a preview. The native macOS menu exposes Create Table under Data. The shortcut only runs with grid focus and does not intercept cell editing.
- An explicit rectangle is used exactly; a single cell suggests its current region. Additional selections and active sorting/filtering are refused. Name and local A1 range are editable; header normalization and record/column counts are previewed. Cancel changes nothing.
- This authoring slice treats the first row as headers, including header-only Tables. Inserting a new header row for headerless data remains part of the structural-history slice.
- Selecting a Table cell shows its name, exact range, record count, Rename, Resize, Banded Rows, and Convert to Range. Convert has a reviewable confirmation. Header tint, alternating body rows, and the active Table outline render only in the viewport; explicit/conditional fills retain precedence, and no per-cell formatting is stamped.
- Editing a header cell invokes a schema rename and rewrites dependent formulas. Invalid edits leave the original header intact and report the reason. Bulk paste/fill/clear/cut, transforms, and Replace All that include headers are refused before mutation. Fill Down may use a header as its source when all destinations are body cells. Multi-header schema paste remains a follow-up.
- Creation, rename, resize, banding, conversion, and header edits use sparse `TableCommit` history entries. Rewind can replay them and locate their Table range. Stale top-level undo/redo reports an error and retains its history position.
- Worksheet sort/AutoFilter refuses Table-bearing sheets pending Table-aware views. Merge refuses Table cells before clearing any values. Desktop Excel export refuses Tables until the user converts them to ranges or saves in `.sheet` format; XLSX Table interchange is still unimplemented.

Automatic growth, calculated-column propagation, Table-backed pivot sources, headerless creation, multi-header paste, and Table-local filters are not part of this desktop slice.

## Next slices

1. Table-aware append and structural undo, headerless creation, calculated-column rules and visible exceptions.
2. Multi-header schema paste and remaining formula editing/interchange integrations.
3. Table-backed pivot sources with stable field IDs, explicit refresh, and stale-state feedback.
4. XLSX interoperability subset, required-feature file compatibility, web/cloud preservation, and export-loss messaging.
5. Local table views and visible-record paste after the current row-view constraints are addressed.

Existing PivotTables remain a separate feature.

## Verification

Behavior tests are in `crates/engine/tests/tables.rs`, `crates/engine/tests/structured_tables.rs`, and `crates/io/tests/tables.rs`; session-host and operation-plan tests cover atomic header rejection. Run:

```sh
cargo test -p visigrid-engine -p visigrid-io -p visigrid-session-host
```

The structured-reference regressions cover syntax/escaping, row context, cross-sheet range consumers, incremental value edits, shape changes, rename/column swaps, destructive edits and stale undo, conversion, empty Tables, cycles, copy/fill, structural moves, formula grouping, default Table names, native/JSON reopening, and semantic fingerprints. Desktop UI and XLSX table interoperability are not exercised by this engine slice.

Verified 2026-10-01: 1,279 tests passed, zero failed, 24 existing tests ignored across these packages. This includes 16 structured-reference integration tests and two new IO regressions, in addition to the foundation's 20 Tables tests.

Desktop validation, 2026-10-01: `cargo test -p visigrid-gpui --bin visigrid` passed 577 tests, zero failures, three existing ignores; `cargo build -p visigrid-gpui --bin visigrid` passed. New tests cover dialog ranges, canonical-row header preflight, schema/formula/style history replay and stale refusal, and structured-selector editor recognition. Linux live checks cover creation preview, controls and pointer geometry, Table/header rename, direct structured formula entry, header rename undo/redo, resize, banding undo, atomic clear refusal, conversion confirmation/cancellation A1 conversion undo, and native save/reopen preserving the Table and structured formula. macOS/Windows UI and XLSX interchange remain untested.
