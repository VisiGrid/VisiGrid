# Print and PDF implementation

Status: basic PDF export is wired into the desktop app on `feat/print-pdf`.
Native printer submission and the visual page preview are still pending.

## Export a PDF

Choose **File → Export PDF…** (also available in the command palette).
Select Active sheet or Selected range, A4/Letter/Legal, portrait/landscape,
Fit columns/Actual size/Fit sheet, and optional page numbers. The dialog reports
page count, scale, and smallest cell text size before opening the native save
prompt. Use a filename ending in `.pdf`. A successful export keeps the receipt
visible, including clipping, small-text, and missing-font notices, with an
explicit **Open PDF** button.

- The adapter captures current calculated display values, dimensions, effective
  conditional formats, agent roles, merges, and the active sort/filter/hide view.
  Screen zoom is not an input. Export uses current cached calculations (including
  manual-calculation mode), does not force volatile recalculation, and does not
  change the workbook path or mark it saved.
- Automatic bounds use nonempty displayed values, excluding covered/hidden cells,
  empty formula results, and distant formatting. Select a rectangle to include
  intentional blank formatting. Intersecting merges expand the selected scope.
  Merges made discontinuous by sorting fail with a clear explanation.
- Capture is synchronous; encoding, shaping, and disk writes run in the background.
  Cancel discards the render and checks between pages/cells. Saving uses a synced
  temporary sibling followed by atomic replacement. Rendering errors cannot
  truncate an existing destination.
- The optional `pdf` feature connects the engine adapter to Krilla and cosmic-text.
  Advanced shaping handles bidi and font fallback; resolved fonts are embedded in
  searchable vector output. The app's bundled IBM Plex Sans faces are always
  available. Other scripts depend on locally installed fallback fonts; missing
  glyphs fail instead of silently producing tofu. Missing requested families are
  reported. Text wraps within existing row heights; export never resizes the sheet.
- Explicit fills, font styling, alignment, borders, wrapped text, merges, and
  left-aligned overflow are rendered. Semantic styles use a light paper palette;
  desktop selection and theme backgrounds are omitted.

This is an incremental export milestone, not the full print specification.
Margins are fixed at 0.5 inches; there is no visual page preview, native Print
command, workbook-wide output, saved print area, persistent settings, page-range
control, gridline/headings controls, or repeated-title controls in the dialog.
The core already supports repeated titles, but that UI remains to be built.
Center-across-selection fidelity and shared grid/PDF text metrics still need work.
External edits after capture do not change the export; reopen the dialog to
capture them. The sheet/range choice captures a fresh snapshot.

Headless QA uses the same adapter and renderer (not a shipped CLI command):

```sh
cargo run -p visigrid-print --features pdf --example export_pdf -- INPUT.xlsx OUTPUT.pdf 'Invoice'
cargo test -p visigrid-print --all-features
```

Linux workbook verification (2026-09-30): all 23 print tests, 83 cell-format
tests, and 132 XLSX tests pass (8 existing XLSX tests ignored). The original
`print-pdf-test.xlsx` exports Invoice / Long report / Wide table / Edge cases
as 1 / 4 / 1 / 1 A4 pages. Poppler confirms embedded, subsetted, Unicode-mapped
fonts; extracted output preserves all 120 unique transaction IDs and the invoice,
report, and annual totals. Hidden markers and distant blank-formatted cells do
not appear. Accented Latin, Greek, Japanese, Arabic, wrapping, merged regions,
and all seven PDF pages were visually checked. Invoice has no clipping warnings;
Edge cases correctly identifies intentional clipping at B16. The long report's
narrow date/description columns produce clipping notices; the wide portrait
table reports text below 8 pt at 53.3% scale. These notices do not resize cells.

The fixture also exposed existing import/display bugs fixed alongside export:
namespace-prefixed XLSX formatting/layout and relationship attributes now parse,
typed ISO dates import as numeric dates, and the Excel `;;;` number format hides
numeric values. Focused regressions cover these cases.

Desktop smoke test on Linux: File → Export PDF opens the settings dialog;
Tab/Space changes paper size, Enter opens the native save prompt, and exporting
Invoice produces a one-page Letter PDF with the correct date and $1,128.24 total.
The success receipt stays open, and cancelling a second native save prompt
returns to the settings without overwriting the file. The workbook remains an
XLSX import. `cargo clippy --no-deps -p visigrid-print --all-features --all-targets
-- -D warnings` passes; existing dependency warnings remain outside this crate.

The product proposal and research live in the planning repository:

- [Print/PDF proposal](../../../planning/visigrid/features/print-pdf-spec.md)
- [Excel complaint research](../../../planning/visigrid/features/print-pdf-excel-research.md)

## Implemented contract

`visigrid-print` has no default dependencies and no GPUI, workbook, printer, or
filesystem dependency. It paginates a caller-supplied immutable visible layout.

- Dimensions are physical points. Unzoomed grid dimensions convert at 0.75
  points per logical unit. Screen DPI, zoom, and default printer are not inputs.
- Source coordinates are retained independently of visible row/column positions.
  The caller supplies visibility and display order; hidden positions are absent.
- Letter, Legal, A4, portrait/landscape, nonnegative margins, and an optional
  fixed 18-point footer reservation determine the page body.
- Actual size and custom scale (10–400%) paginate in both dimensions. Fit columns
  and Fit sheet shrink only, with a 10% lower bound and an explicit limiting
  dimension. Fits below that limit fail rather than silently violating the fit.
- Breaks occur between whole rows/columns and never cross a merge. Merge spans
  combine transitively on an axis. An indivisible group larger than the available
  body fails with its visible positions instead of clipping or looping.
- Repeated titles are leading visible axis counts. They consume space on every
  page and occur only once per page. A title boundary cannot cross a merge, and
  titles must leave both a body position and physical space for it.
- Pages are ordered down then over. Exact boundaries do not add trailing pages.
- A page references bands rather than duplicating cell data. `PagePlan::rect`
  resolves cell/merge positions, including repeated titles, for any renderer.
- Readability reports actual scaled font size and unique cells below 8 pt. The
  threshold is a proposed usability heuristic, not an accessibility standard.
- Inputs are validated and output is bounded to one million visible cell
  positions and one thousand pages. Invalid/empty inputs return typed errors.

The caller is responsible for resolving print scope, range expansion at merges,
calculation state, effective formatting, font selection, and merge geometry
after filtering/sorting. The optional `snapshot` module implements the first workbook adapter; the pure
layout API itself does not resolve these inputs.
Only nonempty displayed text origins belong in `LayoutInput::text`.

## PDF backend spike

The optional `pdf-spike` feature enables Krilla 0.8.2 (MIT/Apache-2.0) for the
`pdf_spike` example. It generates a two-page styled fixture from the same page
plan, including a merged title, repeated headers, borders, Unicode text, footer,
and the app's existing IBM Plex fonts. The production `pdf` feature now uses Krilla with externally shaped runs;
`pdf-spike` remains an example-only feature.

Backend finding: Krilla's `draw_text` convenience API explicitly does not perform
bidi resolution or font fallback and supports only a single script per call.
The fixture therefore exercises Latin text with accented letters and currency
symbols. The production renderer uses cosmic-text shaped runs through `draw_glyphs`;
this original spike alone is not Unicode coverage QA.

```sh
cargo test -p visigrid-print --all-features
cargo clippy -p visigrid-print --all-features --all-targets -- -D warnings
cargo run -p visigrid-print --features pdf-spike --example pdf_spike -- /tmp/visigrid-print-spike.pdf
pdfinfo /tmp/visigrid-print-spike.pdf
pdffonts /tmp/visigrid-print-spike.pdf
pdftotext -layout /tmp/visigrid-print-spike.pdf -
pdftoppm -scale-to 1200 -png /tmp/visigrid-print-spike.pdf /tmp/visigrid-print-spike
```

Use a new output path for each prototype run: the example refuses to overwrite
an existing file. Its direct fixture writer is not the production atomic-export
implementation.

Initial Linux verification (2026-09-29): 17 pagination tests pass with and
without the optional backend. The fixture exports two Letter pages; Poppler
reports both fonts embedded, subsetted, and Unicode-mapped. Text extraction
contains each of the 58 report items once, the title/header twice, accented text,
currency symbols, and both page numbers. Both rendered pages were visually
inspected for clipping, spacing, and header/footer placement. This does not
qualify macOS, Windows, native printers, or arbitrary workbook rendering.

## Remaining print roadmap

1. A common shaped drawing representation for visual preview and PDF; current
   output shares geometry with the page plan but is not a GPUI page preview.
2. Preview navigation, range/title editing, custom margins/scale, printed headings
   and gridlines, page-range controls, and source-revision refresh notices.
3. Exact shared formatting semantics for all grid cases (including center across
   selection and extent growth from text spill), richer clipping locations and
   blank-page diagnostics, and repeat-title controls.
4. Persist settings with native-format migration and structural range tracking;
   preserve them through CLI/headless saves and exclude them from semantic hashes.
5. Platform QA for font resolution, bidi/CJK output, save dialogs, cancellation,
   and replacement on macOS and Windows. Linux test results do not qualify them.

Native printing requires separate Linux/macOS/Windows adapter spikes and real
printer QA. The current feature saves PDFs; it does not submit print jobs.
