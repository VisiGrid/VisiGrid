# Print and PDF implementation

Status: initial foundation and rendering spike on `feat/print-pdf`. This is not
yet a user-facing print or PDF feature. The desktop app has no new menu commands.

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
after filtering/sorting. These are not implemented by accepting a `LayoutInput`.
Only nonempty displayed text origins belong in `LayoutInput::text`.

## PDF backend spike

The optional `pdf-spike` feature enables Krilla 0.8.2 (MIT/Apache-2.0) for the
`pdf_spike` example. It generates a two-page styled fixture from the same page
plan, including a merged title, repeated headers, borders, Unicode text, footer,
and the app's existing IBM Plex fonts. Krilla is a candidate, not a final backend
decision. The feature is not enabled by or linked into the desktop app.

Backend finding: Krilla's `draw_text` convenience API explicitly does not perform
bidi resolution or font fallback and supports only a single script per call.
The fixture therefore exercises Latin text with accented letters and currency
symbols. Production text must use externally shaped runs through `draw_glyphs`
or another proven shaping integration; this spike is not Unicode coverage QA.

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

## Remaining before a PDF milestone

1. Capture active sheet/selection from the real engine and UI at one revision,
   including current filter/sort state, automatic bounds, and partial merges.
2. Extract shared formatting and text layout. Prove wrapping, alignment, spill,
   border ownership, conditional/semantic styles, font fallback, CJK and bidi.
   The spike draws simple text; it does not yet prove these production cases.
3. Add a common shaped drawing representation for preview and output. The
   current plan shares geometry only; it does not guarantee preview/PDF text
   metrics until both renderers consume the same shaped text.
4. Add GPUI preview/settings, title/range editing, print-area explanations,
   clipped-text/blank-page diagnostics, and feasible readability alternatives.
5. Persist settings with native-format migration and structural range tracking;
   preserve them through CLI/headless saves and exclude them from semantic hashes.
6. Add asynchronous generation, cancellation, atomic saving, and snapshot
   revision handling. Validate PDF text extraction and rendered pages on all
   supported platforms with controlled font assets.

Native printing requires separate Linux/macOS/Windows adapter spikes and real
printer QA. Nothing in this initial crate submits print jobs.
