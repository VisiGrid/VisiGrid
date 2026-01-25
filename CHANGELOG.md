# Changelog

## 0.2.9

*This release positions VisiGrid as the Explainable Spreadsheet.*

### Named Range Intelligence (Pro)

Named ranges now integrate with the Inspector and Quick Open.

- **Detail panel**: Select a named range in the Names tab to see value preview, usage count, depth, and verification status.
- **Quick Open**: Press `Ctrl+P` to see named ranges alongside commands (or type `$` to filter to names only).
- **DAG trace**: Single-click a named range to highlight its cells and precedents in the grid. Double-click to jump.

### README Refresh

The README now leads with VisiGrid's category claim: **The Explainable Spreadsheet**. Updated feature lists emphasize trust, causality, and provenance over generic spreadsheet features.

### Demo Workbook

New `fixtures/explainability-demo.csv` demonstrates dependency chains, Verified Mode, and the Inspector in action.

## 0.2.8

*This release completes VisiGrid's explainability loop: verified values, visible dependencies, and provable changes.*

### Explainability: Provenance History (Pro)

VisiGrid now records *why* the grid changed — not just *that* it changed.

- **New History tab** shows a git-log-style list of your latest actions (Paste, Fill, Sort, Clear, Multi-edit).
- **Selecting an entry** reveals a read-only Lua snippet that describes the action in deterministic A1 notation.
- **Copy the Lua snippet** to share, document, or audit how the sheet was produced.
- **Filter box** to quickly find relevant actions by label or scope.

**Pro feature:** Lua provenance + Copy. Free users see the list and can upgrade to view/export provenance.

**Shortcut:** `Ctrl+Shift+Y` opens History.

### Also in this release

- Theme refresh
- Formula bar enhancements

## 0.2.7

- Pro Inspector trust-focused redesign

## 0.2.6

- Reliability & data-safety release

## 0.2.5

- AutoFilter
- Sort

## 0.2.4

- macOS polish
- Freeze panes fix

## 0.2.3

- Multi-color formula references
- F1 inspector peek

## 0.2.2

- Cell borders
- Paste Values (`Ctrl+Shift+V`)

## 0.2.1

- Fill handle
- Background colors

## 0.2.0

- Cross-sheet references
- Zoom controls
- 107 formula functions
