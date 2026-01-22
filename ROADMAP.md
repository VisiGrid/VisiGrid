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
- Clipboard (copy/cut/paste) with formula reference adjustment
  - Relative references shift on paste (A1 → B1)
  - Absolute references stay fixed ($A$1)
  - Mixed references partial shift ($A1, A$1)
- Fill down (Ctrl+D) and fill right (Ctrl+R)

### Navigation & Workflow
- Keyboard-driven navigation (arrows, Ctrl+arrows, Home/End, Page Up/Down)
- Command palette (Ctrl+Shift+P) with fuzzy search (prefixes: `>` commands, `@` cells, `$` named ranges, `:` goto)
- Go To dialog (Ctrl+G)
- Find and Replace (Ctrl+F / Ctrl+H)
  - Incremental search in text and formula cells
  - One dialog, two modes (Ctrl+F/H toggles while open)
  - Token-aware formula replacement (preserves cell references)
  - Replace Next (Enter), Replace All (Ctrl+Enter)
  - Single undo point for Replace All
- Keyboard hints (press `g` for Vimium-style jumping)
- Vim mode (hjkl navigation, optional)
- Session restore (files, scroll position, selection, panels)
- Zen mode (F11)
- Freeze panes (View menu or command palette)

### File Formats
- Native `.sheet` format (SQLite-based, preserves everything)
- CSV/TSV import and export
- JSON export
- Excel import (.xlsx, .xls, .xlsb, .ods) with background processing
- Excel export (.xlsx) with trust-hardening:
  - Formula export with fallback to values for unsupported functions
  - Large numbers (16+ digits) as text to preserve precision
  - Export Report dialog shows exactly what changed

### Formula Engine (107 functions)
- Math, logical, text, lookup, date/time, statistical, array
- Dynamic arrays with spill behavior
- SUMIF/SUMIFS, COUNTIF/COUNTIFS, AVERAGEIF/AVERAGEIFS
- XLOOKUP, FILTER, SORT, UNIQUE
- Financial: PMT, FV, PV, NPV, IRR
- Autocomplete with signature help
- Syntax highlighting and error validation
- Context help (F1)

### Multi-Sheet & Organization
- Multiple sheets per workbook
- Sheet tabs (add, rename, delete, reorder)
- Cross-sheet references (=Sheet2!A1)
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
- macOS (Universal binary, signed and notarized, Homebrew)
- Windows (x64)
- Linux (x86_64, tar.gz, AppImage, Homebrew, AUR)

### View
- Zoom (50%-200%, Ctrl+Shift+=/-, Ctrl+Mousewheel, persisted per session)

### CLI (v0.1.8+)
- Headless spreadsheet operations via `visigrid-cli`
- `calc`: Evaluate formulas against piped data (CSV, TSV, JSON, lines)
- `convert`: Transform between file formats
- `list-functions`: Show all supported functions
- Column references (A:A), headers support, spill output
- Typed JSON output (numbers/booleans preserved)
- 3ms cold start, pipe-friendly, deterministic
- See `visigrid-cli --help` for usage

### Fill Handle (v0.2.1+)
- Drag the corner handle on the active cell to fill down or right
- Formula references adjust correctly (relative/absolute)
- Axis locks automatically (vertical or horizontal, no diagonals)
- Single undo step per fill operation

### Cell Background Colors (v0.2.1+)
- Apply background colors from Format menu or command palette
- 9-color palette + clear
- Live preview while cell is selected
- XLSX export: colors written to `.xlsx` files (patternFill)
- Note: XLSX import does not read fills (calamine library limitation)

---

## In Progress

None currently.

---

## Planned

### 0.2.2 — Borders v1

Simple borders that round-trip with XLSX. No complexity explosion.

**Scope (intentionally limited):**
- Presets only: None / All / Outline
- Solid style, 1px width, black color
- No per-edge customization in v1

**Engine:**
- [ ] Add `borders: CellBorders` to `CellFormat`
- [ ] Store per-cell edge intents (top/right/bottom/left)
- [ ] Deterministic conflict rule: if either cell defines edge, draw it

**XLSX I/O:**
- [ ] Export: write `<border>` elements
- [ ] Import: read border styles

**UI:**
- [ ] 3-button preset picker (None / All / Outline)
- [ ] Apply to selection with undo

**Deferred to later:**
- Per-edge width/style/color
- Full border picker UI
- Excel-perfect "last writer wins" (needs write-order tracking)

### Near-term
- Paste Special (values, formulas, transpose, operations)
- Series Fill (smart pattern detection)
  - Numbers: 1, 2, 3 → 4, 5, 6
  - Dates: Jan 1 → Jan 2, Jan 3...
  - Months: Jan, Feb → Mar, Apr...
  - Weekdays: Mon, Tue → Wed, Thu...
  - Custom step sizes
- Data validation (dropdowns, constraints)
- Conditional formatting (basic rules)
- Comments/notes on cells
- Print to PDF

### Medium-term
- Sparklines (=SPARKLINE formula, Unicode mini-charts)
- AutoFilter and Sort
- Cell context menu (right-click)
- Merged cells
- Split view (Ctrl+\)
- Problems panel (Ctrl+Shift+M) - all formula errors

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
| gpui/Zed | Nested submenu support | Cross-platform nested menus for context menus and menu bar. See [#19837](https://github.com/zed-industries/zed/issues/19837) |

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
