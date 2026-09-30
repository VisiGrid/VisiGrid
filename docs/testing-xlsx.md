# Headless XLSX testing

The CLI and desktop share `visigrid-io`. Run the ordinary regression tests without building a GUI:

```sh
cargo test --locked -p visigrid-io
```

The `Headless I/O fidelity` workflow runs this on Linux, macOS, and Windows. The downloaded external corpus is an exploratory local audit; it is not yet a compatibility gate in CI.

## External corpus audit

`fixtures/xlsx-corpus.json` pins six Apache POI workbooks by upstream revision and SHA-256. Download each `url` to a source directory using its `file` name. The manifest links the upstream license; original binaries remain local under ignored `test-results/`, rather than being redistributed in the repository.

```sh
cargo build --locked -p visigrid-io --example xlsx_probe
# Adds the controlled basics workbook to the source directory:
target/debug/examples/xlsx_probe --generate test-results/xlsx-corpus/source
python3 scripts/xlsx-audit.py test-results/xlsx-corpus/source test-results/xlsx-corpus/results
```

On Windows use `target/debug/examples/xlsx_probe.exe` and the installed Python command. Optionally copy the two `fixtures/issue17-*.xlsx` files into the source directory. No Excel or Python packages are required.

Each workbook takes three paths: unchanged export, A1 text edit and export, and native `.sheet` save/reload/export (including layout). Every exported file is reopened by VisiGrid. Python independently checks ZIP CRCs, XML parsing, cell types/values/formula representations, sheet names/order/visibility, merge ranges, freeze panes, counts of conditional rules and validations, defined names, and chart/table/drawing parts. The controlled basics fixture additionally checks explicit RGB colors, bold, borders, number formats, sizes, hidden rows/columns, and ordinary/cross-sheet formula calculations.

The audit writes `report.json`, exported workbooks, and import/export diagnostics. Nonzero exit means an execution, structural, checksum, or controlled-format assertion failed. Semantic differences are recorded for investigation, not automatically accepted as compatibility passes. Known feature losses need reviewed per-feature expectations before this corpus becomes a CI gate.

Limitations: shared formulas are reported as representation changes instead of being normalized; cached formula values are not compared; rule counts do not prove rule equivalence; general style/theme normalization and relationship-target validation are not implemented. XML parsing does not prove that Excel will open a workbook without repair. The original corpus contains unsupported and intentionally troublesome features. Keep an Excel visual/opening check before release.

## Initial findings

Nine workbooks and 27 paths were exercised on macOS. Two native persistence bugs were fixed, with regression coverage for all three workbook save APIs and the legacy text loader:

- Numeric-looking or formula-looking literal text was reinterpreted on reload.
- Merge ranges were saved only for the active sheet and loaded onto sheet zero. Per-sheet merge metadata now preserves all sheets; old files without this metadata retain the legacy reader behavior.

Remaining observed limitations: Boolean cells export as text; 13 three-dimensional references in `55906-MultiSheetRefs.xlsx` export as `#ERR` text; chart objects, table definitions, defined names, and the conditional-formatting sample's two rules are lost. The chart fixture's six shared formulas correctly expand to `D1*D1` through `D6*D6`, though the raw representation audit reports them as differences. Formula cached results in XLSX exports are zero until recalculated by a consuming application.

The controlled workbook's styles pass the machine checks. Review the source and exported `basics.xlsx` side by side for appearance, plus the issue17 formatting/frozen-overlay workbooks in the desktop app. Headless testing does not exercise native grid rendering.

## Desktop visual follow-up

Comparing the original basics workbook in Excel and VisiGrid exposed three rendering bugs, despite preserved file contents:

- Selecting the blue header suppressed its explicit white text color.
- A merge crossing the frozen column repeated its text and extended too far because its viewport-relative origin was clamped to zero. Offscreen origins now remain negative for clipping.
- The cell below/right of a merge repeated the merge overlay's border. Identical shared edges now have one painter; stronger independent adjacent borders remain visible.

The corrected QA app was checked on macOS. Row 1 remained fixed while scrolling down to row 61, and column A remained fixed while navigating horizontally to column U. Frozen panes were already supported and imported; their divider is subtle. Geometry and border-ownership regression tests accompany the fixes.

F2 follow-up: editing `='Other sheet'!A1+B3` shaded B3 correctly but drew its dashed outline over A2. The parser already omitted the other sheet's A1; the outline used scrolling-only coordinates. Reference outlines now use the grid's four frozen/scrolling regions and clipping, preserving offscreen range origins. The exact formula and frozen-reference geometry have regression tests (46 formula-related tests passed). Verified in the Formula QA app: the only local outline encloses B3.

Sequential-open fix: opening `WithVariousData.xlsx` after `basics.xlsx` retained hidden column E, hidden row 13, and previous sizes. Imported layout application now clears all four workbook-specific maps before applying the new workbook's layout. Verified in Import QA: default widths/height restored, E and row 13 visible. Scroll origins also respect imported frozen panes after an in-window import.

The C4 hyperlink's underline was correctly imported but dropped by the overflowing-text layer. That layer now applies its stored underline flag; verified visually in Import QA. This fixes appearance, not hyperlink-target preservation. The General-format numbers still display an extra trailing zero in VisiGrid. At the time of this initial check, comments and print header/footer settings were not imported. Excel Notes import/export is now covered by `testing-comments.md`; print header/footer import remains unsupported. C9/C10 are ordinary labels, while the separate print header reads “This is the header on sheet 1” and the footer expands the worksheet name and page number. Confirmed the header in Excel's Page Layout view.
