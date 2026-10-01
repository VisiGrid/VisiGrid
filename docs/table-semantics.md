# Tables: engine foundation

Status: first implementation slice, 2026-10-01. This is an engine API and persistence contract, not a released desktop Tables feature. The product plan lives in the Obsidian notes “VisiGrid Tables Spec” and “VisiGrid Tables Research”.

## Model

A `DataTable` identifies an editable rectangle of existing sheet cells. It owns schema, not a second copy of its data. The inclusive rectangle starts with one header row; a header-only table is valid. `TableId` is workbook-unique. `TableColumnId` is scoped to a table. IDs survive renaming, native/JSON saves, and undo/redo. Allocator high-water marks prevent reuse after removal, shrinking, or undo.

Table names share a case-insensitive namespace with named ranges. This slice accepts conservative ASCII names and rejects reference-like names. Column names are case-insensitive, nonempty text; blank/duplicate headers receive deterministic names. Original valid names are reserved before generating suffixes: `Amount, Amount, Amount2` becomes `Amount, Amount3, Amount2`. Names starting with `=` or containing controls are normalized too. Creation previews normalization without mutation.

Table ranges cannot overlap another table, a merged region, pivot output, or an existing array spill. Dynamic arrays cannot spill into a Table, including blank body cells. Ordinary scalar formulas in the body remain ordinary formulas.

## Operations and undo

`Workbook::create_table`, `rename_table`, `rename_table_columns`, `resize_table`, and `remove_table` validate before changing state and return an opaque `TableCommit`. `apply_table_commit(commit, undo)` replays that commit with schema/header preconditions. A stale commit fails before mutation. No body snapshot is stored; undo retains schema, header preconditions, and the original typed values of changed headers.

Creation uses an explicit rectangle whose first row is already the header. It converts normalized headers to text, preserving their explicit formatting. Header insertion and automatic region detection are future UI work.

Resize keeps the top-left corner fixed and changes the bottom/right edges. Surviving columns retain IDs, new columns get fresh IDs, and shrinking leaves released cell values intact. Remove deletes metadata and leaves values/formulas/explicit formats intact. The `banded_rows` flag is persisted but is not rendered by this slice.

Headers change through the schema API. Low-level cell setters refuse direct header value writes; tracked workbook setters return no recalculation delta for a refused write. Session batches and operation plans reject header writes during preflight, so they cannot report a successful partial edit. Hosts must use this preflight before other multi-cell editing flows are exposed.

Structural edits entirely before a Table move its bounds with its cells. Edits after it leave its bounds unchanged. Insertions within a Table and deletions intersecting it are temporarily refused. Use explicit resize/remove for schema changes. This keeps the existing structural undo contract intact until table-aware row/column history is integrated. Table-bearing sheet duplication also temporarily refuses instead of duplicating IDs. Sheet removal/restoration updates name reservations and rejects conflicting restoration.

## Persistence

Native `.sheet` workbook saves (ordinary, metadata, full) store a strict versioned `tables` metadata catalog. The single-sheet native API routes table-bearing sheets through the workbook format. Invalid metadata, duplicate identities/names, invalid bounds, ownership collisions, and headers that disagree with schema cause load failure. Missing catalog in an old native file means no tables. Allocators persist even after the last table is removed.

Full JSON emits version 3 only when a workbook/sheet has Table history, with a required `table_catalog`. Ordinary exports retain existing v1/v2 output. Older JSON readers reject v3 rather than silently dropping definitions.

**Release gate:** old native readers may ignore this new metadata and drop it when saving. A minimum-reader/required-feature compatibility policy is still needed before public Tables authoring ships. XLSX/ODS/CSV and web/cloud editing do not yet preserve table definitions. Native semantic fingerprints also need to include Table semantics before structured formulas or verification rely on them.

## Next slices

1. Structured-reference parser, resolver, rename rewriting, dependency invalidation, and copy/fill semantics.
2. Desktop creation preview, header editing, table styling/context commands, and history integration; all paste/fill/clear paths must preflight ownership atomically.
3. Table-aware append and structural undo, calculated-column rules and visible exceptions.
4. Table-backed pivot sources with stable field IDs, explicit refresh, and stale-state feedback.
5. XLSX interoperability subset, required-feature file compatibility, web/cloud preservation, and export-loss messaging.
6. Local table views and visible-record paste after the current row-view constraints are addressed.

No new keyboard shortcut or desktop command is exposed by this foundation. Existing PivotTables remain a separate feature.

## Verification

Behavior tests are in `crates/engine/tests/tables.rs` and `crates/io/tests/tables.rs`; session-host and operation-plan tests cover atomic header rejection. Run:

```sh
cargo test -p visigrid-engine -p visigrid-io -p visigrid-session-host
```

Verified 2026-10-01: 1,261 tests passed, zero failed, 24 existing tests ignored across these packages (including 20 new Tables regression tests). Desktop UI and XLSX table interoperability are not exercised by this foundation.
