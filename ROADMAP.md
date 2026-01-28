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

### Formula Engine (108 functions)
- Math, logical, text, lookup, date/time, statistical, array
- Dynamic arrays with spill behavior
- SUMIF/SUMIFS, COUNTIF/COUNTIFS, AVERAGEIF/AVERAGEIFS
- XLOOKUP, FILTER, SORT, UNIQUE, SPARKLINE
- Financial: PMT, FV, PV, NPV, IRR
- Autocomplete with signature help
- Syntax highlighting and error validation
- Context help (F1): hold to peek cell inspector
  - Formula cells: formula text, computed value, format, precedents, dependents
  - Value cells: value, format, dependents
  - Range selection: count, sum, average, min, max
  - Positioned near selection

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

### Cell Borders (v0.2.2+)
- All Borders, Outline Borders, Clear Borders presets
- Canonicalized border storage (shared edges handled correctly)
- Adjacency-based rendering (no doubled lines)
- XLSX export preserves borders
- See: [docs/features/borders-spec.md](docs/features/borders-spec.md)
- **Deferred:** Border colors, Medium/Thick UI, per-edge controls, XLSX import

### Paste Values (v0.2.2+)
- Ctrl+Shift+V pastes computed values instead of formulas
- Internal clipboard stores typed `Value` objects (not display strings)
- Clipboard metadata ID for reliable internal vs external detection
- External clipboard: leading zeros preserved, `=` prefix becomes text
- Edit mode: inserts top-left cell only, canonical text (no scientific notation)
- See: [docs/features/paste-values-spec.md](docs/features/paste-values-spec.md)
- **Deferred:** Paste Special dialog, Paste Formats, Paste Transpose

### Sparklines (v0.2.3+)
- `=SPARKLINE(range)` formula returns Unicode mini-bar-chart
- `=SPARKLINE(range, "winloss")` for win/loss charts (▲▼▬)
- 8-level bar heights using Unicode block characters (▁▂▃▄▅▆▇█)
- Auto-scales to data min/max

### Fill Handle Improvements (v0.2.3+)
- Excel-style rendering: solid fill with contrasting border
- Shows at bottom-right corner of range selections (not just single cells)
- Smaller, more precise hit target (6px visual, 14px hit area)

### macOS Keyboard (v0.2.3+)
- Modifier key preference: choose Cmd (default) or Ctrl for shortcuts
- Setting: `keyboard.modifierStyle` in `~/.config/visigrid/settings.json`
- Delete key now clears selected cells (maps to backspace on Mac keyboards)
- System shortcuts (Cmd+Q, Cmd+W, Cmd+,) always use Cmd regardless of preference

### Multi-Color Formula Reference Highlighting (v0.2.3+)
- Grid: referenced cells/ranges highlighted with 8 rotating colors
- Stable colors per reference (duplicate refs share one color)
- Selection always visually dominant over formula highlights
- Formula bar: reference tokens colored to match grid highlights
- Live updates while editing formulas
- Unicode-safe: token spans handle non-ASCII correctly (char→byte fixed)
- See: [docs/features/formula-reference-highlighting-spec.md](docs/features/formula-reference-highlighting-spec.md)
- **Deferred:** Marching ants animation (Phase 3)

### macOS Transparent Titlebar (v0.2.4+)
- Zed-style chrome blending with `appears_transparent: true`
- Traffic lights positioned inward at (9, 9)
- Custom titlebar with document identity:
  - Primary: filename + dirty indicator (12px, full contrast)
  - Secondary: provenance text (10px, muted, quieter)
- 34px titlebar height (matches Zed)
- Double-click to zoom (native macOS behavior)
- Draggable title bar area via `WindowControlArea::Drag`
- Subtle chrome scrim (8px gradient fade into content)
- Hairline border separator (50% opacity)
- macOS handles inactive window dimming natively
- See: [docs/features/title-bar-and-menus.md](docs/features/title-bar-and-menus.md)

### Alt Menu Accelerators (v0.2.4+, macOS)
- Excel-style Alt+letter shortcuts open scoped Command Palette
- Alt+F (File), Alt+E (Edit), Alt+V (View), Alt+O (Format), Alt+D (Data), Alt+H (Home/Format)
- Opt-in via Preferences (disabled by default to preserve Option key)
- Scope badge shown in palette header when filtered
- Backspace clears scope when query is empty
- Never intercepts keys during cell/formula editing
- See: [docs/features/alt-accelerators.md](docs/features/alt-accelerators.md)

### Default App Prompt (v0.2.4+, macOS)
- "Make default" title bar chip for non-native file types
- Scoped to specific file type: "Open .csv files with VisiGrid"
- Right-aligned, Zed-style banner (muted, smaller than filename)
- Guardrails prevent annoying behavior:
  - Only after successful file load (not during import errors)
  - Not for temp files, unsaved documents, or native .vgrid
  - Session cool-down (once per session per file type)
  - 7-day cool-down if user ignores (doesn't dismiss or act)
  - Permanent dismiss via ✕
- Post-click feedback:
  - Success: "Default set" for 2 seconds
  - Needs completion: "Finish in System Settings" + "Open"
- Uses `duti` when available, falls back to System Settings
- See: [docs/features/title-bar-and-menus.md](docs/features/title-bar-and-menus.md)

### Data Validation (v0.3.4+)
- Dropdown lists (comma-separated or range reference)
- Numeric constraints (between, greater than, less than, equals)
- Invalid cell highlighting (red circle marker)
- Jump to next invalid (F8)
- Validation exclusions (exempt cells from rules)
- Command palette integration
- See: [docs/features/data-validation-spec.md](docs/features/data-validation-spec.md)

### AI Integration (v0.3.5+)
- **Ask AI** (`Ctrl+Shift+A`) — natural language to formula proposals
- **Provenance tracking** — every AI-inserted cell tagged with `MutationSource::Ai`
- **Explain Differences** — right-click history → diff report with AI-touched filter
- **Explain This Change** — 2-4 sentence descriptions of individual changes
- **AI Summary** — optional, manual-trigger summaries of change sets
- **Trust model**: AI describes but never modifies from audit context
- See: [docs/features/future/ai-reconciliation-spec.md](docs/features/future/ai-reconciliation-spec.md)

---

## Planned

### Near-term
- Paste Special expansion (formulas, transpose, operations)
- Series Fill (smart pattern detection)
  - Numbers: 1, 2, 3 → 4, 5, 6
  - Dates: Jan 1 → Jan 2, Jan 3...
  - Months: Jan, Feb → Mar, Apr...
  - Weekdays: Mon, Tue → Wed, Thu...
  - Custom step sizes
- Conditional formatting (basic rules)
- Comments/notes on cells
- Print to PDF

### Medium-term
- Cell context menu (right-click)
- Merged cells
- Problems panel (Ctrl+Shift+M) - all formula errors

### Explainability Initiative

Make causality, impact, and recomputation visible. VisiGrid becomes an **Explainable Spreadsheet**.

See full spec: [docs/features/explainability-roadmap.md](docs/features/explainability-roadmap.md)

**Phase 1 — Dependency Graph Foundation**
- Persistent dependency graph (precedents/dependents in O(1))
- Cross-sheet dependency tracking
- Ordered recompute engine with cycle detection at edit-time
- See: [docs/features/dependency-graph-spec.md](docs/features/dependency-graph-spec.md)

**Phase 2 — Verified Mode**
- User-facing toggle with status bar promise: "Verified ✓ — values are current"
- Summary stats (recalc time, cells, depth)

**Phase 3 — Inspector Extensions**
- Dependency depth, recompute timestamps, evaluation traces
- Mini DAG visualization (fan-in/fan-out)
- Basic Inspector becomes free; deep explainability stays Pro

**Phase 4+ — Provenance & Beyond**
- Lua snippets for structural operations (fill, sort, transform)
- Named ranges as first-class inspector entities
- Blast radius preview before large operations

### Polish
- Windows title bar integration (custom titlebar with integrated menu)
  - Eliminates "menu on 2nd line" non-native feel
  - See [docs/features/windows-titlebar-spec.md](docs/features/windows-titlebar-spec.md)
  - Windows-first, behind feature flag, with acceptance checklist

### Long-term: Systems of Record

Connect to authoritative data sources without sync risk.

Import data from ERPs, payment processors, and ledgers as read-only snapshots. Every pull is versioned with a timestamp. User controls when to refresh. No silent syncs. No write-back.

Target sources:
- Stripe / payment processors
- QuickBooks / accounting systems
- Bank feeds
- Database exports

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
