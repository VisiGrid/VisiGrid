# Tables: engine, structured references, desktop authoring, row growth, and calculated columns

Status: desktop authoring, safe row growth, and calculated columns implementation, 2026-10-01. This is the staged implementation contract, not a declaration that the complete Tables release is ready. The product plan lives in the Obsidian notes “VisiGrid Tables Spec” and “VisiGrid Tables Research”.

## Model

A `DataTable` identifies an editable rectangle of existing sheet cells. It owns schema, not a second copy of its data. The inclusive rectangle starts with one header row; a header-only table is valid. `TableId` is workbook-unique. `TableColumnId` is scoped to a table. IDs survive renaming, native/JSON saves, and undo/redo. Allocator high-water marks prevent reuse after removal, shrinking, or undo.

Table names share a case-insensitive namespace with named ranges. This slice accepts conservative ASCII names and rejects reference-like names. Column names are case-insensitive, nonempty text; blank/duplicate headers receive deterministic names. Original valid names are reserved before generating suffixes: `Amount, Amount, Amount2` becomes `Amount, Amount3, Amount2`. Names starting with `=` or containing controls are normalized too. Creation previews normalization without mutation.

Table ranges cannot overlap another table, a merged region, pivot output, or an existing array spill. Dynamic arrays cannot spill into a Table, including blank body cells. Ordinary scalar formulas in the body remain ordinary formulas.

## Operations and undo

`Workbook::create_table`, `rename_table`, `rename_table_columns`, `resize_table`, and `remove_table` validate before changing state and return an opaque `TableCommit`. `apply_table_commit(commit, undo)` replays that commit with schema/header and dependent-formula preconditions. A stale commit fails before mutation. No whole-body snapshot is stored; undo retains schema, header preconditions, original typed values of changed headers, sparse append writes, and sparse formula-source changes. New dependent formulas that require additional rewrites make replay stale. Creation also captures existing dependent sources so undo restores their originally unbound state.

Creation uses an explicit rectangle. With **My data has headers** enabled (the default), its first row supplies normalized text headers and keeps its explicit formatting. With headers disabled, every selected row remains data: one whole worksheet row is inserted at the selection's top boundary, and `Column1`, `Column2`, … become the new headers. Cells in all columns at or below that row move down, including adjacent data and Tables; A1 references, calculated rules and structural metadata follow the existing row operation. The preview shows the final range and record count and explicitly describes the whole-row movement. A single-cell selection still suggests its current region.

Headerless creation returns a single `TableCommit` and increments the workbook revision once. It stages row insertion and creation on a temporary copy-on-write workbook, publishing only after both succeed. Undo/redo also stage both operations together: a late collision or spill cannot leave a partial insertion/removal. History retains schema, inserted-row preconditions, structural formula rewrites and print setup, without storing a workbook/body snapshot. The desktop shifts row heights and hidden-row positions and restores them on undo; rewind uses the same engine commit. Generic row insertion still governs freeze-pane positions.

Headerless preview refuses overlapping Tables, merges, spills and pivot output, structural protection violations and insufficient space for the extra row. A last-row value, comment or explicit format also prevents insertion; the desktop additionally checks last-row height/visibility metadata. Final creation validates again after row movement, catching formulas that start spilling at their new coordinates. Stale undo refuses changes to the inserted row or captured formula sources, and failed replay retains the original workbook and history position.

Resize keeps the top-left corner fixed and changes the bottom/right edges. Surviving columns retain IDs, new columns get fresh IDs, and shrinking leaves released cell values intact. Remove converts dependent structured formulas to absolute A1 references and deletes metadata, preserving values and explicit formats. Conversion refuses referenced empty bodies or unresolved references because they have no lossless A1 representation. `set_table_style` changes persisted banding through the same guarded commit API.

Headers change through the schema API. Low-level cell setters refuse direct header value writes; tracked workbook setters return no recalculation delta for a refused write. Session batches and operation plans reject header writes during preflight, so they cannot report a successful partial edit. Hosts must use this preflight before other multi-cell editing flows are exposed.

Structural edits entirely before a Table move its bounds with its cells. Edits after it leave its bounds unchanged. Whole worksheet rows inserted within the body grow the Table; deleted body rows shrink it, including deletion of every record to leave a header-only Table. Generic deletion of the header row remains refused; worksheet column edits follow the contract below. `TableRowHistory` stores before/after schema bounds alongside ordinary row history; undo restores exact bounds even when reinserting at the bottom of a now-empty Table. Replay checks the current Table catalog before mutation. The desktop continues to retain deleted cells, comments, row heights, print setup and formula rewrites in the same row history entry. This helper is not a standalone cell snapshot. Table-bearing sheet duplication also temporarily refuses instead of duplicating IDs. Sheet removal/restoration updates name reservations and rejects conflicting restoration. Removing a sheet with Table references from other sheets is refused until sheet history can capture those rewrites.

## Safe append

`Workbook::append_table_rows` accepts an explicit row count and sparse body writes, and returns one guarded `TableCommit`. Ordinary workbook setters and file loading never infer append intent. Appending validates the entire new stripe, including untouched columns: existing values/comments, another Table, merges, spills, pivot output, or the grid boundary refuse the operation before any write. Explicit Resize remains the way to include existing records. Formatting-only empty cells are allowed and retain their formatting. Append changes membership without inserting worksheet rows or shifting adjacent data.

Desktop entry points:

- Nonempty typing immediately below a Table, within its width, appends one row when the new stripe is empty.
- A rectangular paste starting in the body or immediately below it, contained within its width and extending below the bottom, appends through the pasted last row. Represented blank records count. Existing body writes and new bounds share one undo entry.
- A paste crossing both side and bottom boundaries refuses with Resize guidance. Writes separated by a blank row or entirely to the right do not imply growth. Single-cell paste broadcast/fill across a selected range retains existing fill behavior.
- Tab from the last body cell appends one empty row and selects its first column. An in-progress edit in that last cell joins the append commit. A header-only Table uses the **Add row** control; that control is also available for nonempty Tables.
- Clearing cell values retains membership. Appending and structural row edits refuse active sorting/filtering until cleared.

Normal paste, Paste Values and Paste Formulas share append preflight. An internal normal paste carrying merges/comments (or replacing destination comments) refuses growth with guidance to use Paste Values or resize first; it never silently drops those objects. Omitted cells in calculated columns receive their rule; supplied values/formulas, including explicit blank cells, take priority.

Undo/redo checks both schema and typed values owned by the append commit. Redo refuses newly occupied space; undo refuses edits to appended cells or new data/comments in previously untouched appended cells. Rewind replays append and whole-row history through the same engine operations. Grouped arbitrary structural mutations are not exposed as a new API in this slice.

## Worksheet column edits

Whole worksheet columns can be inserted or deleted with Tables and calculated columns present. Inserting strictly inside a Table expands its schema with empty columns, unique case-insensitive `ColumnN` headers and freshly allocated IDs. Inserting at/before the first column moves the Table; inserting after the last column leaves it unchanged. Use Resize to expand at an outer edge.

Deleting part of a Table removes those fields and shrinks the bounds; surviving IDs, names, calculated rules, values and overrides follow their cells. Deleting every column of any affected Table is refused before mutation: convert that Table to a range first. Grid-edge and pivot collision preflights still apply. A1 references shift structurally; structured references to removed fields become permanent `#REF!`, including rules on other sheets and header-only rules. Reusing a removed name does not repair those references. A structured span loses its reference if an endpoint is deleted; deleting an interior field contracts the span between its surviving endpoints.

`TableColumnHistory` retains before/after schemas and affected rule sources, alongside ordinary sparse column history for deleted values, comments, formatting, widths, print setup and formula rewrites. Undo restores exact field IDs and formulas; redo reuses the same inserted IDs and headers. Allocation high-water marks never roll back. Rewind uses the same engine path. Replay validates the expected schemas/rules before mutation; rejected desktop undo/redo keeps its history position, and rejected deletions leave widths unchanged. This is schema metadata rather than a second copy of Table records.

## Calculated columns

A column may own a formula rule plus its authored row offset. The origin is retained rather than rebasing the stored formula to the first record, which could lose relative references above row 1. Formulas are materialized in the ordinary cell store; existing relative/absolute A1 translation applies, and structured tokens stay symbolic. Import, creation and bulk setters never infer a rule.

- Direct formula entry in an otherwise empty body column creates its rule and fills all records as one guarded Table commit. Tab with a last-cell edit and typing an appended record can establish the rule within the same append commit.
- In an already populated column, direct entry changes only that cell. Selecting a formula offers **Use formula for entire column** with a preview of the existing values/formulas it will replace.
- In an established column, individual cell edits are overrides. A small amber corner and the contextual **Override** label identify them. **Restore column formula** restores the selected cell in one undoable operation.
- **Edit column formula** updates cells that follow the rule and preserves overrides. **Replace entire column** is separate and previews the overwrite before applying. All these operations support undo/redo and rewind.
- New records from Add row, Tab, typing, append paste or whole-row insertion receive the rule. Explicit paste writes take priority; explicit blank cells stay blank, while omitted trailing cells receive formulas. Clearing a formula does not remove the rule.

Exception state is derived from authoritative cell contents, rather than duplicated in a mutable exception registry. A cell is an exception when its value/formula differs from the projected rule; an empty cell is an exception. Equivalent parsed formulas (including the same formula pasted back) follow the rule again. This deliberate implementation choice makes existing clear/paste/fill, history and persistence paths consistent and avoids stale override flags. It does not retain provenance for an explicit override that is identical to the rule.

Rules follow Table/column renames, conversion, removed-column references and whole-row history, including external rules referring to an edited sheet. Deleting all records retains the rule for future rows. Worksheet-column history retains rules and their authored origins, including deleted calculated fields and rules on other sheets.

Rule creation rejects malformed formulas and formulas that evaluate to arrays at preflight. Later inputs can still make a formula return an array; the existing Table spill barrier reports `#SPILL!` without writing outside the cell. Cycles retain the engine's existing behavior. Schema/rule commits check the cells they will overwrite before replay, so stale undo cannot silently erase later overrides.

Native/JSON catalogs containing rules use Tables metadata version 2, which the previous Table reader rejects. Version 1 remains valid for workbooks without rules. Formula source/origin are persisted and included in semantic fingerprints, even for header-only Tables; overridden cell values remain ordinary persisted cells. The broader minimum-native-reader release gate still applies.

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
- **My data has headers** defaults on, including header-only Tables. Turning it off inserts a header row and preserves all selected records. Tab/Shift+Tab includes the checkbox; Space toggles it, Enter applies, and Escape cancels without mutation.
- The creation dialog pairs name/source range fields and previews up to three columns and two records in a small grid, with total counts and the resulting Table range. Headerless creation has a separate row-insertion explanation. Adjusted column names remain visible beneath the preview, and invalid ranges show an inline error. The primary action is visually distinct from Cancel.
- Selecting a Table cell shows its name, exact range, record count, Add row, Rename, Resize, Banded Rows, and Convert to Range. Convert has a reviewable confirmation. Header tint, alternating body rows, and the active Table outline render only in the viewport; explicit/conditional fills retain precedence, and no per-cell formatting is stamped.
- Editing a header cell invokes a schema rename and rewrites dependent formulas. Invalid edits leave the original header intact and report the reason. Bulk paste/fill/clear/cut, transforms, and Replace All that include headers are refused before mutation. Fill Down may use a header as its source when all destinations are body cells. Multi-header schema paste remains a follow-up.
- Creation, rename, resize, banding, conversion, and header edits use sparse `TableCommit` history entries. Rewind can replay them and locate their Table range. Stale top-level undo/redo reports an error and retains its history position.
- Worksheet sort/AutoFilter refuses Table-bearing sheets pending Table-aware views. Merge refuses Table cells before clearing any values. Desktop Excel export refuses Tables until the user converts them to ranges or saves in `.sheet` format; XLSX Table interchange is still unimplemented.

Table-backed pivot sources, multi-header paste, and Table-local filters remain outside this slice.

## Next slices

1. Multi-header schema paste and remaining formula editing/interchange integrations.
2. Table-backed pivot sources with stable field IDs, explicit refresh, and stale-state feedback.
3. XLSX interoperability subset, required-feature file compatibility, web/cloud preservation, and export-loss messaging.
4. Local table views and visible-record paste after the current row-view constraints are addressed.

Existing PivotTables remain a separate feature.

## Verification

Behavior tests are in `crates/engine/tests/tables.rs`, `crates/engine/tests/table_growth.rs`, `crates/engine/tests/structured_tables.rs`, and `crates/io/tests/tables.rs`; session-host and operation-plan tests cover atomic header rejection. Run:

```sh
cargo test -p visigrid-engine -p visigrid-io -p visigrid-session-host
```

The structured-reference regressions cover syntax/escaping, row context, cross-sheet range consumers, incremental value edits, shape changes, rename/column swaps, destructive edits and stale undo, conversion, empty Tables, cycles, copy/fill, structural moves, formula grouping, default Table names, native/JSON reopening, and semantic fingerprints. Desktop UI and XLSX table interoperability are not exercised by this engine slice.

Verified 2026-10-01: 1,279 tests passed, zero failed, 24 existing tests ignored across these packages. This includes 16 structured-reference integration tests and two new IO regressions, in addition to the foundation's 20 Tables tests.

Desktop validation, 2026-10-01: `cargo test -p visigrid-gpui --bin visigrid` passed 577 tests, zero failures, three existing ignores; `cargo build -p visigrid-gpui --bin visigrid` passed. New tests cover dialog ranges, canonical-row header preflight, schema/formula/style history replay and stale refusal, and structured-selector editor recognition. Linux live checks cover creation preview, controls and pointer geometry, Table/header rename, direct structured formula entry, header rename undo/redo, resize, banding undo, atomic clear refusal, conversion confirmation/cancellation A1 conversion undo, and native save/reopen preserving the Table and structured formula. macOS/Windows UI and XLSX interchange remain untested.

Row growth validation, 2026-10-01: the full engine/IO/session-host run passed 1,287 tests (24 existing ignores), and the additional native/JSON append round-trip test passed. The final desktop suite passed 578 tests (3 existing ignores); the launchable build passed. Eight new engine regressions cover explicit intent, 1,000-row paste with blank records, collisions/stale replay, one-revision append, header-only membership, grid bounds and structural row history. Desktop rewind covers append and whole-row insertion/deletion. Linux live checks confirmed typing growth with undo/redo, Tab with and without an in-progress edit, Add row including header-only Tables, overlapping rectangular paste with one-step undo, deleting all body rows with undo/redo, and occupied-row refusal with unchanged bounds/data. The temporary QA workbook was saved. macOS/Windows UI remain untested.

Calculated-column validation, 2026-10-01: the engine/IO/session-host suite passed 1,301 tests (24 existing ignores), and the desktop suite passed 579 tests (3 existing ignores), with zero failures. The launchable desktop build passed. Twelve engine regressions in `crates/engine/tests/calculated_columns.rs` cover automatic fill, populated-column opt-in, overrides, explicit blanks versus omitted cells, append/Tab commits, authored A1 origins, schema changes, structural history and guarded replay. Native/JSON round trips preserve rules and overrides; desktop rewind preserves cleared overrides across rule changes. Linux live checks confirmed automatic fill with one-step undo/redo, visible overrides, Add row filling, editing the rule while preserving an override, and Restore column formula. The temporary workbook was saved. macOS/Windows UI and XLSX Table interchange remain untested.

Column-history validation, 2026-10-01: the engine/IO/session-host suite passed 1,310 tests (24 existing ignores), and the desktop suite passed 580 tests (3 existing ignores), with zero failures. The launchable build passed. Eight regressions in `crates/engine/tests/table_columns.rs` cover internal/boundary insertions, clipped deletions, original field identity on undo/redo, allocation high-water marks, case-insensitive generated headers, calculated rules and overrides, permanent deleted-field references, external/header-only rules, multi-Table edits and atomic refusal. Native/JSON round trips preserve schema and allocator state; desktop rewind matches live schema, values and dependent formulas. Linux live checks confirmed insertion, referenced-column deletion and undo/redo. Saved-file inspection verified exact field IDs, the restored formula rule, destructive `#REF!` on redo and the preserved `999` override. macOS/Windows UI remain untested.

Headerless-creation validation, 2026-10-01: the engine/IO/session-host suite passed 1,318 tests (24 existing ignores), and the desktop suite passed 581 tests (3 existing ignores), with zero failures. The launchable build passed. Seven regressions in `crates/engine/tests/headerless_tables.rs` cover generated headers, preserved records/formats/comments, cross-sheet references, neighboring calculated Tables, blank records, bottom-edge and protected-region refusal, late-spill rollback, stale replay and formerly unbound structured references. Native/JSON round trips and desktop rewind preserve all records. Linux live checks confirmed default-on preview, keyboard toggling, cancellation, creation and single-step undo/redo. Saved-file inspection verified the same Table/column identities on redo, adjusted formulas, adjacent-cell movement and restored row height. macOS/Windows UI remain untested.
