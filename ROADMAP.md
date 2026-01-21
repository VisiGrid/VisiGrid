# VisiGrid Roadmap

What's built, what's next, and what's not planned.

> Current status: Early access (pre-v1.0). Expect breaking changes to `.sheet` format.

---

## Shipped

### Core Spreadsheet
- GPU-accelerated grid rendering (GPUI)
- Cell editing with formula bar
- Selection, multi-selection, and range operations
- Multi-edit (type once, apply to all selected cells with formula shifting)
- Undo/redo
- Clipboard (copy/cut/paste)
- Paste Special (values, formulas, transpose, operations)
- Fill down (Ctrl+D) and fill right (Ctrl+R)

### Navigation & Workflow
- Keyboard-driven navigation (arrows, Ctrl+arrows, Home/End, Page Up/Down)
- Command palette (Ctrl+Shift+P) with fuzzy search (prefixes: `>` commands, `@` cells, `$` named ranges, `:` goto)
- Go To dialog (Ctrl+G)
- Find with incremental search (Ctrl+F)
- Keyboard hints (press `g` for Vimium-style jumping)
- Vim mode (hjkl navigation, optional)
- Session restore (files, scroll position, selection, panels)
- Zen mode (F11)

### File Formats
- Native `.sheet` format (SQLite-based, preserves everything)
- CSV/TSV import and export
- JSON export
- Excel import (.xlsx, .xls, .xlsb, .ods) with background processing

### Formula Engine (97 functions)
- Math, logical, text, lookup, date/time, statistical, array
- Dynamic arrays with spill behavior
- SUMIFS, COUNTIFS, XLOOKUP, FILTER, SORT, UNIQUE
- Autocomplete with signature help
- Syntax highlighting and error validation
- Context help (F1)

### Multi-Sheet & Organization
- Multiple sheets per workbook
- Sheet tabs (add, rename, delete, reorder)
- Named ranges with full IDE support:
  - Create (Ctrl+Shift+N)
  - Go to definition (F12)
  - Find all references (Shift+F12)
  - Rename across all formulas (Ctrl+Shift+R)

### Formatting
- Bold, italic, underline, strikethrough (bold/italic not rendering on Linux, see [docs/font-rendering-issue.md](docs/font-rendering-issue.md))
- Number formats (currency, percent, decimal, general)
- Format Cells dialog (Ctrl+1)
- Cell alignment
- Column/row resize (drag or double-click to auto-fit)

### Developer Features
- Lua scripting console (Ctrl+Shift+L)
- Inspector panel (Ctrl+Shift+I) - precedents, dependents, diagnostics
- Configurable keybindings (`~/.config/visigrid/keybindings.json`)
- Themes (10+ built-in, Ctrl+K Ctrl+T)
- URL/path detection (Ctrl+Enter to open)

### Platform
- macOS (Universal binary, signed and notarized)
- Windows (x64)
- Linux (x86_64, tar.gz and AppImage)

---

## In Progress

### Polish
- Freeze panes (lock rows/columns while scrolling)

### Formula Coverage
- AVERAGEIF, AVERAGEIFS
- Financial: PMT, FV, PV, NPV, IRR

---

## Planned

### Near-term
- Split view (Ctrl+\)
- Problems panel (Ctrl+Shift+M) - all formula errors
- Find and replace (Ctrl+H)
- Cross-sheet references (=Sheet2!A1)
- Zoom (Ctrl++/-)
- XLSX export
- Data validation (dropdowns, constraints)
- Conditional formatting (basic rules)
- Comments/notes on cells
- Print to PDF

### Medium-term
- CLI / headless mode (pipe-friendly, `visigrid calc`, `visigrid convert`)
- Sparklines (=SPARKLINE formula, Unicode mini-charts)
- AutoFilter and Sort
- Cell context menu (right-click)
- Fill Handle (drag to extend selection pattern)
- Series Fill (smart pattern detection: 1,2,3 → 4,5,6)
- Merged cells

### Long-term: Systems of Record

Connect to authoritative data sources without sync risk.

Import data from ERPs, payment processors, and ledgers as read-only snapshots. Every pull is versioned with a timestamp. User controls when to refresh. No silent syncs. No write-back.

Target sources:
- Stripe / payment processors
- QuickBooks / accounting systems
- Bank feeds
- Database exports

### Long-term: AI Reconciliation

AI helps you reason about data without modifying it.

AI proposes operations (compare, flag, explain). User reviews and approves. VisiGrid executes deterministically. Every step visible and auditable.

**AI can:** Compare records, explain differences, flag anomalies, summarize discrepancies

**AI cannot:** Edit cells without approval, fetch data directly, hide steps, invent values

### Long-term: Extensibility
- Plugin architecture (WASM-based)
- Custom functions
- Data connectors
- Minimap for large sheets

---

## Upstream Contributions

Fixes we may contribute to dependencies:

| Project | Issue | Details |
|---------|-------|---------|
| gpui/Zed | Linux font rendering | Bold, italic, per-cell fonts don't render. See [docs/font-rendering-issue.md](docs/font-rendering-issue.md) |

---

## Non-Goals

Explicit rejections, not "not yet":

| Feature | Rationale |
|---------|-----------|
| VBA/macro compatibility | Lua scripting instead |
| Real-time collaboration | Local-first philosophy |
| Web version | Desktop-native performance |
| Mobile/tablet | Desktop-first workflows |
| Pivot tables (v1) | Complexity; later if demand |
| Charts (v1) | Separate concern; later |
| Perfect Excel formatting | Diminishing returns |

---

## Versioning

VisiGrid uses [SemVer](https://semver.org/): `MAJOR.MINOR.PATCH`

Current: Early access (0.x). Breaking changes possible before v1.0.
