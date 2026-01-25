# Changelog

## 0.2.9

VisiGrid now completes the explainability loop: you can verify results, inspect structure, and export the intent behind changes.

### Provenance History (Pro)

VisiGrid records *why* the grid changed — not just *that* it changed.

- **History tab** shows a git-log-style list of high-level actions (Paste, Fill, Sort, Clear, Multi-edit).
- **Select an entry** to view a canonical, read-only Lua snippet describing the action in deterministic A1 notation.
- **Copy the Lua snippet** to share, document, or audit how a sheet was produced.
- **Filter history** by label or scope.

**Pro feature:** View + Copy Lua provenance. Free users see the history list and can upgrade to export provenance.

**Shortcut:** `Ctrl+Shift+Y` opens History.

### Named Range Intelligence (Pro)

Named ranges are now first-class model entry points.

- **Named range details** show value preview, depth, verified status, and usage count.
- **Quick Open** (`Ctrl+P`) includes named ranges by default.
- **One-click tracing** highlights dependencies for the selected named range.

### Also

- Theme refresh (contrast + readability improvements)
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
