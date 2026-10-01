# Parquet export

`vgrid convert` exports one worksheet as a typed Parquet table. The shared
implementation lives in `visigrid-io::parquet_export` and is also used by the
production web API through the CLI.

```sh
vgrid convert report.xlsx --sheet Data -t parquet --headers -o data.parquet
vgrid convert report.sheet -t parquet --headers --parquet-plan
vgrid convert report.sheet -t parquet --headers --text-column Status -o data.parquet
vgrid convert input.csv -t parquet --headers --where 'Amount>0' --select 'ID,Amount' > data.parquet
```

`--headers` uses the first populated row as column names and excludes it and
preceding empty rows from the data. Without it, column names are A, B, C, …
and all rows are data. Blank and duplicate names (case-insensitive) are refused.
`--text-column NAME` is repeatable and matches the output name, after `--rename`.
Filtering and selection happen before analysis, using the existing CLI filter
semantics. Hidden rows and columns are included unless explicitly filtered.

`--parquet-plan` writes JSON to stdout, with `ready`, `row_count`, and per-column
`data_type`, `null_count`, `issue_count`, bounded cell-level `issues`, and
`warnings`. A valid analysis request exits successfully even when `ready` is
false; consumers must inspect that field. Malformed options or headers fail
normally. It cannot be combined with `-o` and never writes a Parquet file.

## Value contract

Every selected row is checked; inference is not based on a sample. All columns
are nullable. Numeric data comes from the engine's typed values, never from
rounded display strings.

| Current sheet values | Parquet type |
| --- | --- |
| Text, including numeric-looking text and literal `TRUE` | STRING |
| Whole numbers within the exact integer range ±(2^53 − 1) | INT64 |
| Fractional numbers, larger numeric values, or whole/fractional mixtures | DOUBLE |
| Computed boolean results | BOOLEAN |
| Numbers consistently formatted as calendar dates | DATE |
| Numbers consistently formatted as times of day | TIME(MICROS, local) |
| Date-times, or a mixture of dates and date-times | TIMESTAMP(MICROS, local) |
| Entirely null column / header-only table | Nullable STRING, with a notice |

Built-in date/time formats and Excel custom formats are recognized. Elapsed
formats such as `[h]:mm:ss` remain numbers. Numeric sections of a custom
date/time format must agree on their temporal type; otherwise simplify the
format or explicitly export as text. Dates with hidden fractional days are
refused rather than truncating their time. The fictitious Excel date
1900-02-29 is refused. Calendar dates are limited to years 0001–9999; serial
zero maps to 1899-12-31. Time of day must be >= 0 and < 1 day. Timestamps and
times round to microseconds; no UTC timezone is invented.

Mixed incompatible types block export and identify the offending cells.
`--text-column` converts that column's underlying values to strings, including
raw numeric date serials; it is not a formatted-display export. Nulls remain
null and actual empty text remains an empty string. Formula errors, missing
computed results, and non-finite numbers block export even with a text override.
Formulas are exported as their computed values, not as expressions. Inputs
loaded through the CLI are recalculated by their existing loaders; supported
custom-function fallbacks may carry previously computed results.

This exports **current spreadsheet values**, not an exact original-file round
trip. Import may already have converted booleans to text, normalised text,
reduced timestamp precision, or represented large integers and decimals as text.
The exporter does not recover source schemas, timezone provenance, exact decimal
arithmetic, binary values, or nested structures. Currency formatting does not
create a DECIMAL column. CSV has no type schema; its import inference happens
before export and before `--text-column`.

The writer uses Snappy compression and row groups of 2,048 rows. Path output is staged in the target
directory and replaces its destination only after successful completion. Binary
stdout is spooled first and refused on an interactive terminal.

## Web integration

The production Rails API exposes authenticated
`POST /api/sheets/:id/export_parquet`, requiring the same sheet read access as
XLSX export. JSON body:

```json
{
  "document": {"format":"visigrid-json","version":2,"sheets":[{"name":"Data","cells":[]}]},
  "sheet_index": 0,
  "headers": true,
  "text_columns": [],
  "plan": true
}
```

With `plan: true` it returns the CLI analysis JSON. Otherwise it returns an
`application/vnd.apache.parquet` attachment. The browser holds a frozen snapshot
while the export dialog is open, checks the schema, and lets the user mark
columns as text. It includes unsaved changes without saving or publishing them.
The server recalculates formulas on each request; volatile results can change
between checking and downloading, and the writer validates again.

The initial web limits are 30 MB input/output, two million cells across workbook
bounding rectangles, a 1 GB converter memory budget, and a 20-second conversion
timeout. CLI export has no such
web limits. No object is uploaded to R2. Responses are private/no-store.

This requires deploying both the API and a CLI build with the export options,
then the web build. It works through a browser on desktop, iPad and iPhone; it is
not an offline browser/WASM implementation or a native iOS CLI. The isolated
Loco pilot and public share-link page do not expose this action yet. Native
desktop/TUI export menus are a separate integration; the shared writer and CLI
are available on supported native build targets.
