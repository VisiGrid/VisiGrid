# DuckDB files

VisiGrid opens local `.duckdb` files read-only. Each base table becomes a
worksheet, with column names in row 1. The desktop app, CLI, terminal viewer,
and production web import use the same native reader. No separate DuckDB
installation is needed.

## Desktop import

Opening a database shows a table picker before importing any records. Select
table names to preview eight records and the source column types; checkboxes
choose the tables to import. The preview shows at most twelve columns, while
import includes every column. Search filters the table list without changing
the selected tables. Empty tables can be imported as column headers.

The picker shows exact row counts and blocks tables that exceed worksheet or
import limits. The combined selection must also fit the 10-million-cell budget.
Metadata, previews, and the selected-table import share a read-only transaction
on a background reader thread. Cancel leaves the current workbook unchanged;
late responses are discarded. An in-flight query may finish before the worker
releases the database, so cancelling is not a guarantee of immediate lock release.

Use Up/Down to preview tables, Space to select, Tab/Shift+Tab to move between
controls, Ctrl/Cmd+F to search, Ctrl/Cmd+Enter to import, and Escape to cancel.
Imported tables become an editable, unsaved workbook. Saving uses Save As;
edits never write back to DuckDB. If the current workbook has unsaved edits,
opening a DuckDB file creates a separate window.

## CLI and terminal viewer

```sh
# Browse table tabs in the terminal (5,000 records per table by default)
vgrid peek warehouse.duckdb
vgrid peek warehouse.duckdb --sheet main.orders --max-rows 100
vgrid peek warehouse.duckdb --sheet orders --json --max-rows 10
vgrid peek warehouse.duckdb --shape

# Keep every table as a worksheet
vgrid convert warehouse.duckdb -t json-full -o workbook.json
vgrid convert warehouse.duckdb -t sheet -o workbook.sheet
vgrid convert warehouse.duckdb -t xlsx -o workbook.xlsx

# Select one table before reading its records
vgrid convert warehouse.duckdb --sheet main.orders -t csv -o orders.csv

# Export one worksheet into a NEW database containing main.data
vgrid convert workbook.xlsx --sheet Orders -t duckdb --headers --export-plan
vgrid convert workbook.xlsx --sheet Orders -t duckdb --headers -o orders.duckdb
vgrid convert mixed.csv -t duckdb --headers --text-column Amount -o mixed.duckdb
```

Tables are ordered by schema, then table name. Names normally appear as
`main.orders`; unusual names are quoted. `convert` and `peek` accept a displayed
name, a zero-based index, or an unqualified table name if unique across schemas.
Without `--sheet`, single-table outputs use the first table; workbook outputs
keep all tables. `peek --json` emits the selected table, defaulting to the first.
`sheet inspect` and `sheet import` also recognize `.duckdb` and accept displayed
tab names or indices.

## Data and limits

This imports table values, not a live database connection. Refresh, arbitrary
SQL, views, macros, indexes, constraints, and database writeback are outside this
format's contract. The source database is never changed. Close other writers and
checkpoint the database before copying or uploading it; the main file alone may
not include changes still in a `.wal` file. The embedded engine is DuckDB 1.5.5;
files requiring a newer storage version may need an older-compatible export.

Reads use a single transaction. Preview record order follows DuckDB's scan order,
not a promised sort order. Counts describe the whole table; previews explicitly
report omitted rows. There is no paging or fetching more rows while scrolling.
The default preview cell budget is 10 million across all loaded tables.
`--max-rows 0` requires `--force` above 200,000 rows. Even with `--force`, worksheet
dimensions remain a hard limit: 1,048,575 data rows plus a header and 16,384 columns.

Full imports/conversions refuse to truncate and use a 10-million-cell aggregate
budget (including headers). Select a table with the CLI or split the source if
it does not fit. The web upload limit remains 30 MB; compressed database size is
not the same as the expanded worksheet size.

DuckDB tables are transferred through temporary Parquet files and the shared
Parquet importer. Large integers that exceed exact spreadsheet numeric precision
become text; strings are never executed as formulas. Spreadsheet storage does
not retain the original database schema, all decimal/temporal precision, or
trailing all-null records as database records. Some complex types may be
unsupported. This is not a lossless database backup or round trip.

Exports follow the [Parquet conversion policy](parquet-export.md): validate
every output row, reject incompatible mixed types and formula errors, and allow
explicit text conversion. They contain calculated values, not formulas or
formatting. One selected sheet becomes `main.data`. Export refuses an existing
destination, including a concurrent writer creating that path; it never replaces
or appends to an existing database. A complete checkpointed file is staged next
to the destination before publication. DuckDB output requires `-o`.

## Web and build integration

The web editor's Export menu offers DuckDB with the same type-review dialog as
Parquet. `/api/sheets/:id/export_duckdb` uses the existing sheet read permission,
CSRF checks, frozen browser snapshot, server-side calculation, and private
download flow. No revision or database connection is saved. iPhone/iPad browsers
use this server-backed flow; the native CLI is not running inside the browser.
The isolated Loco pilot does not expose DuckDB import/export.

Deploy the updated CLI with the Rails API before enabling the updated web client.
The engine is statically bundled through `libduckdb-sys` with its Parquet
extension. Native builds now require a C++ compiler and take longer on a cold
cache. CI limits compilation to two jobs; release builds compile the desktop
app and CLI together so Cargo can share the native build. The WASM formula
engine remains separate. Supported native targets still
need their normal platform release checks.

Extension auto-install and auto-loading, persistent secrets, external access,
and configuration changes are disabled for the connection. Only VisiGrid's
private temporary directory is allowlisted for the Parquet transfer. Imports
enumerate base tables; views are not imported and no SQL query input is exposed.
DuckDB evaluates any generated columns on those base tables.

References: [DuckDB Rust build options](https://duckdb.org/docs/current/clients/rust/overview),
[DuckDB security configuration](https://duckdb.org/docs/current/operations_manual/securing_duckdb/overview),
[storage compatibility](https://duckdb.org/docs/stable/internals/storage),
[checkpointing](https://duckdb.org/docs/current/sql/statements/checkpoint).
