# VisiGrid Roadmap

What's built, what's next, and what's not planned.

> The CLI and deterministic engine are the source of truth. GUI features evolve on top of a stable computation core.

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
- Format Painter (v0.3.6+) — capture cell format, click to apply, Esc to cancel
- Color Picker (v0.3.6+) — theme grid, standard colors, hex input, recent colors, `ColorTarget` architecture
- Clear Formatting (v0.3.6+) — resets all format properties to default
- Format Bar (v0.3.8+) — toggleable toolbar between formula bar and column headers (View → Format Bar). Font family, font size input, B/I/U toggles, fill color, text color, alignment. Tri-state controls for mixed selections. Engine setters for font size and text color with undo/redo.

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
- Headless spreadsheet operations via `vgrid`
- `calc`: Evaluate formulas against piped data (CSV, TSV, JSON, lines)
- `convert`: Transform between file formats
- `list-functions`: Show all supported functions
- Column references (A:A), headers support, spill output
- Typed JSON output (numbers/booleans preserved)
- 3ms cold start, pipe-friendly, deterministic
- See `vgrid --help` for usage

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

### Fill Handle & Series Fill (v0.2.3+, enhanced v0.3.5)
- Excel-style rendering: solid dark fill with contrasting border (10px visual, 18px hit area)
- Corner cap positioning: handle overlaps selection border by 1px inward
- Hover feedback: subtle glow + crosshair cursor
- Shows at bottom-right corner of range selections (not just single cells)
- **Series Fill (v0.3.5):**
  - Single numbers copy by default; hold Ctrl/Cmd for series (1→2→3)
  - Built-in lists auto-extend: months, weekdays, quarters
  - Two+ cell selections detect linear step (1,3→5,7,9)
  - Alphanumeric patterns: Item1→Item2, Row Z→Row AA
  - Leading zeros preserved: 001→002→003
  - Drag preview: border-only (no fill) for clear action feedback
  - First-use tip in status bar
- See: docs/features/done/series-fill-spec.md

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
- See: docs/features/future/ai-reconciliation-spec.md

### Merged Cells (v0.3.7+)
- Full merged cell support across engine, rendering, navigation, clipboard, and UI
- Merge Cells (Ctrl+Shift+M) and Unmerge Cells (Ctrl+Shift+U)
- Data-loss confirmation dialog, overlap guard, contained merge replacement
- Navigation treats merges as atomic data units (arrow, Ctrl+Arrow, Tab, Enter, Go To)
- Copy/paste recreates merges at destination; Cut removes source merges
- See: [docs/features/done/merged-cells-spec.md](docs/features/done/merged-cells-spec.md)

---

## In Progress

Features with significant implementation already shipped.

- **Data Validation** — Dropdown lists, number/date/text constraints, error alerts, XLSX import/export, circle invalid data. Spec: [docs/features/data-validation-spec.md](docs/features/data-validation-spec.md)
- **Explainability** — Dependency graph, verified mode, inspector, provenance, history panel, CLI replay. Spec: [docs/features/explainability-roadmap.md](docs/features/explainability-roadmap.md)
- **Context Menu** — Flat menus complete. Nested submenus blocked on gpui upstream (zed#19837).
- **Ask AI / AI Reconciliation** — Ask AI + Explain Differences working, OpenAI provider.
- **Custom Functions** — Lua scripting exists; needs formula integration and sandbox polish.

---

## Planned

### Near-term

Ready to build, no major blockers.

- **Cell Comments** — Text notes on cells with red triangle indicator, hover preview, edit/delete
- **Print to PDF** — Paginated PDF export with print preview, page setup, headers/footers
- **Conditional Formatting** — Highlight rules, color scales, data bars, icon sets
- **Paste Special Phase 2-3** — Arithmetic paste operations and Transpose
- **Problems Panel** — Bottom panel aggregating all workbook errors with filtering and navigation
- **Merged Cells extensions** — Merge Across (merge each row separately), context menu integration

### Medium-term

Need infrastructure, design decisions, or upstream dependencies.

- **Range Picker** — Excel-style RefEdit control for selecting ranges from modal dialogs
- **Split View** — Horizontal/vertical/four-way panes for comparing distant regions
- **Windows Title Bar** — Custom title bar integrating menu bar. See [docs/features/windows-titlebar-spec.md](docs/features/windows-titlebar-spec.md)

### Long-term

Major infrastructure investment.

- **Plugin Architecture** — WASM-based sandboxed plugins for custom functions, data connectors, UI panels
- **Data Connectors** — SQL databases, REST APIs, GraphQL, S3
- **Systems of Record** — Read-only SaaS integrations (Stripe, QuickBooks, Plaid). OAuth, versioned snapshots
- **Minimap** — Bird's-eye view sidebar showing data density and quick navigation

---

## Upstream Contributions

Fixes we may contribute to dependencies:

| Project | Issue | Details |
|---------|-------|---------|
| gpui/Zed | Linux font rendering | Bold, italic, per-cell fonts don't render on Linux (gpui framework limitation) |
| gpui/Zed | Nested submenu support | Cross-platform nested menus for context menus and menu bar. See [#19837](https://github.com/zed-industries/zed/issues/19837) |
| gpui/Zed | Modifier-only key release events | `KeyUpEvent` not sent when releasing modifier keys (Option/Alt/Ctrl/Cmd) without other keys. Prevents Excel-style double-tap Option for KeyTips. Workaround: Option+Space trigger instead. **Decision: Only implement double-tap Option if gpui adds modifier-only key release events; otherwise KeyTips stays Option+Space.** |

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
