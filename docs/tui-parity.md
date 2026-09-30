# Terminal viewer parity

Assessment: 2026-09-30. The TUI is `vgrid peek`, a read-only viewer in
`crates/cli/src/tui/`. It shares importers and the formula engine with the
desktop, but has its own data adapter and event loop. It is not a terminal
spreadsheet editor.

## This update

- Parquet opens in interactive, plain, JSON, and shape modes. Schema names
  become headers; record 1 is the first data row. The shared Parquet importer
  accepts a preview limit, so peek does not materialize the whole file first.
  Dates and timestamps use ISO strings instead of spreadsheet serial numbers.
- `.vgrid` uses the native workbook reader, alongside `.sheet`.
- `.xls`, `.xlsb`, and `.xlsm` use the same Excel reader as `.xlsx` and `.ods`.
  This reads cell data; it does not execute macros.
- Explicit `--sheet` selection works in plain output. JSON selection accepts
  the same case-insensitive names and zero-based indexes as the TUI.
- `--json --recompute` now recomputes imported Excel/ODS formulas.
- Plain previews show loaded versus total rows. JSON keeps its existing
  `{columns, rows}` envelope and reports truncation on stderr.
- Native workbooks now have the preview cell guard too. Workbook guards use
  the requested preview size, so reducing `--max-rows` actually helps.

Run from `visigrid/app`:

```bash
cargo run -p visigrid-cli -- peek crates/cli/tests/fixtures/parquet_orders.parquet
cargo run -p visigrid-cli -- peek data.parquet --shape
cargo run -p visigrid-cli -- peek data.parquet --json --max-rows 100
cargo run -p visigrid-cli -- peek report.xlsx --sheet Summary --recompute
```

An interactive terminal starts the TUI automatically. Redirected output uses
a plain table. `--plain` forces a table; `--tui` requires terminal input/output.
Arrow keys or hjkl navigate, PgUp/PgDn page, g/G jump to first/last loaded row,
Tab switches workbook sheets, `?` opens help, and q exits.

Finding and summarizing (added after 0.40.0) work on the loaded rows:

- `/` searches every cell (case-insensitive); `n`/`N` step through matches,
  which are highlighted.
- `[` / `]` sort by the cursor column, ascending or descending. Numbers sort
  by value (`1,200`, `$5`, `(5)` included), text case-insensitively, blanks
  last. Sorts are stable, so sort by the tiebreak first. `R` restores file
  order. Row numbers stay file row numbers.
- `F` opens the cursor column's frequency table as a new tab.
- `P` opens a pivot prompt, prefilled from the cursor column:
  `rows=Region column=Month values=sum:Amount`. The result opens as a tab,
  computed by the desktop's pivot engine. Without a header row (CSV opened
  without `--headers`) the first row supplies the field names.
- On a derived tab, `q` returns to the tab it came from.

When the preview is truncated, derived tab names say "(loaded rows)". For a
whole-file answer use `vgrid pivot FILE`, which refuses truncated Parquet
rather than summarizing part of it.

## Current parity

| Capability | Desktop | TUI after this update |
| --- | --- | --- |
| CSV/TSV, Parquet, Excel/ODS, native files | Open/import | Read-only preview through shared importers |
| Multiple workbook sheets | Editable tabs | Tab switching and `--sheet` selection |
| Formula results | Recalculated workbook | Native files recalculate; Excel/ODS default to cached values, opt-in `--recompute` |
| Formula source | Formula bar and editing | Native workbook status line; imported workbook source is not exposed |
| Document settings | Desktop loads sidecar calculation/layout settings | Peek loaders do not apply desktop sidecars |
| Number/date formats | Desktop cell formats | Parquet dates/times use ISO; workbook previews still use raw display values |
| Formatting, comments, merged cells, charts | Desktop rendering | Display strings only; no visual parity |
| Find/go-to, sort/filter | Desktop actions | Find (`/`, `n`/`N`) and single-column sort over loaded rows; no filter or go-to |
| Pivot tables | Field-list drawer, linked to source | `P` pivot and `F` frequency tabs (read-only, loaded rows); `vgrid pivot` for files and sessions |
| Range selection, clipboard, fill | Desktop selection semantics | Single-cell navigation only |
| Editing, formula entry, undo/redo, save | Supported | Not implemented |
| `.duckdb`, SQL table/query browsing | Not implemented | Not implemented |
| Large-file access | Imported into the workbook | Bounded snapshot; no fetching more rows while scrolling |

The broader CLI already has evaluation, conversion, inspection, and scripted
workbook operations. Those commands do not imply interactive TUI editing.

## Limits that matter

- Default: 5,000 rows per sheet. `--max-rows 0` requests all rows; above 200,000
  rows this requires `--force`. Positive explicit limits bypass that row guard.
- Workbook and Parquet preview grids have a 10-million-cell guard unless
  `--force` is supplied. This is per sheet, not a total memory budget.
- Parquet currently passes through the spreadsheet engine: at most 1,048,575
  data records plus its header, and 16,384 columns. `--force` cannot remove
  engine limits. Oversized columns are refused by peek; row truncation is shown.
- Parquet bounds record materialization, but row-group decoding may still do
  more I/O than the requested preview size. `--shape` currently loads a preview.
- Excel and native readers load the workbook before building bounded preview
  grids. CSV reads the source text and scans records to count the total.
  None of these is a database-style paged reader.
- Parquet JSON preserves numeric-looking IDs and represents nulls as null.
  It uses existing engine conversions, not a lossless Parquet interchange
  model: dates are displayed strings, and boolean-like text can be interpreted
  as booleans. CSV and workbook JSON still infer types from display strings.
- `peek` accepts a file path. Direct stdin-to-interactive-viewer support still
  needs separate terminal input handling; existing CLI pipes cover other commands.

## Recommended next steps

1. **Data inspection first:** add find/next, go-to row or column, full-cell
   inspection, formatted workbook values, column selection, and copy/export
   selection. Keep typed export values separate from display formatting. Clearly label
   operations that only cover loaded preview rows.
2. **Shared data-source layer:** provide schema, typed values, row counts,
   cancellable page reads, and query results independent of `Sheet`. Use this
   in the desktop, CLI, and TUI to avoid repeating format dispatch and caps.
   Whole-dataset search/sort/filter should run in this layer.
3. **DuckDB through that layer:** open databases read-only by default, list
   schemas/tables/views, select a table or execute an explicit query, page the
   results, and handle database locks and unsupported values clearly. Expose
   the same source options in all three interfaces. Parquet import alone does
   not provide these capabilities.
4. **Interactive spreadsheet editing, if needed:** hold a live Workbook rather
   than display-string snapshots; share command/history operations; implement
   edit modes, selection, clipboard, undo/redo, and save. Follow
   `selection-semantics.md` and `fill-copy-semantics.md`. Define format fidelity
   and save behavior before allowing edits to imported files.

Before promising identical calculation behavior, also align document settings
and test custom-function availability across builds. Sharing the engine alone
does not guarantee identical host configuration.

For the Parquet/DuckDB audience, prioritize steps 1–3. Full desktop editing
parity is a separate, substantially larger project.
