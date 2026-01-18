# Design Document

A lightweight, native, open-source spreadsheet inspired by Excel 2003's simplicity and modern code editor UX.

---

## Core Philosophy

- **Native performance**: Rust + iced, no web wrappers
- **Keyboard-first**: Excel 2003 shortcuts as default, fully remappable
- **Themeable**: JSON theme files, ship with light/dark/nord/gruvbox
- **Minimal UI**: No ribbon. Menu bar + command palette
- **Opinionated**: Make decisions, don't offer 50 options

---

## Non-Goals (Hard Boundaries)

These are explicit rejections. Do not revisit without a compelling reason.

| Non-Goal | Rationale |
|----------|-----------|
| VBA/macro compatibility | Different paradigm. Lua scripting instead. |
| Perfect Excel formatting parity | Diminishing returns. Support common cases only. |
| Printing engine (v1) | Punt to PDF export → system print. |
| Real-time collaboration (v1) | Architectural complexity. Single-user first. |
| XLSX read/write (v1) | Massive spec. CSV + native format first. |
| Charts (v1) | Separate concern. Export data → external tools. |
| Pivot tables (v1) | Complex feature. Consider v2+. |
| Workbook-level security | Encryption, passwords, etc. Later. |
| Mobile/tablet | Desktop-first. Touch is a different app. |

---

## Performance Contract

Concrete, measurable, tested. Not aspirational.

### Benchmarks

| Metric | Target | Test Method |
|--------|--------|-------------|
| Cold start to interactive | <300ms (Linux), <500ms (Windows) | criterion bench, no file loaded |
| Cold start with last file (10k cells) | <500ms | criterion bench |
| Scroll performance (1M row sheet) | 60fps, visible cells only | perf HUD in debug mode |
| Keystroke to visible change | <16ms median, <50ms p99 | input latency bench |
| Recalc (10k formula cells) | <100ms, non-blocking | criterion bench |
| Memory (empty sheet) | <50MB | runtime measurement |
| Memory (1M cells, numbers only) | <500MB | runtime measurement |
| Binary size (release, stripped) | <30MB | CI check |

### Testing Strategy

- `criterion` benchmarks in CI, fail on regression >10%
- Debug build includes perf HUD (frame time, recalc time, memory)
- Synthetic test sheets: 1k, 10k, 100k, 1M rows
- Input latency measured with instrumented event loop

### Recalc Behavior

- **Never block UI thread** - calculation runs on background thread
- **Cancelable** - new edit cancels in-progress recalc
- **Incremental** - only recalc dirty cells and dependents
- **Progress indicator** - show recalc status for long operations

---

## Input Model v0.1 (Ship This)

Minimum complete spec. Validate in the spike, don't debate edge cases.

### 1. Modes (Explicit, No Ambiguity)

Three modes, period:

| Mode | Focus | Keystrokes Do | Exit Via |
|------|-------|---------------|----------|
| **Navigation** | Grid | Move selection, invoke commands, typing starts edit | Enter edit mode |
| **Edit** | Cell editor | Edit text, arrows move cursor | Esc (cancel) or Enter/Tab (commit) |
| **Command** | Palette | Fuzzy search, execute action | Esc or execute |

Mode transitions:
- Navigation → Edit: Type any character, F2, Enter, or double-click
- Edit → Navigation: Enter (commit), Tab (commit), Esc (cancel), click elsewhere (commit)
- Navigation ↔ Command: Ctrl+Shift+P toggles
- Command always returns to Navigation after executing

### 2. Selection Model

**Data structure:**
```rust
struct Selection {
    ranges: Vec<Range>,      // Ordered list, can be discontiguous
    active_range: usize,     // Index into ranges
    anchor: (usize, usize),  // Anchor cell for extending
}

struct Range {
    start: (usize, usize),   // (row, col), always <= end
    end: (usize, usize),     // Inclusive
}
```

**Invariants:**
- Ranges are normalized: `start.row <= end.row && start.col <= end.col`
- There is always at least one range
- Active cell = `ranges[active_range].start`

**User operations:**

| Action | Result |
|--------|--------|
| Click cell | Single-cell selection, clear others |
| Shift+Click | Extend from anchor to clicked cell |
| Shift+Arrow | Extend active range from anchor |
| Ctrl+Click | Add new single-cell range (discontiguous) |
| Ctrl+Shift+Click | Add new range from anchor |
| Click+Drag | Select rectangle |
| Ctrl+A | Select all cells with data |

### 3. Multi-Edit Semantics (The Heart)

**Key decision: Typing with multi-cell selection REPLACES ALL cells.**

This is our differentiator. Not "edit active only" like Excel.

| Edit Type | Behavior |
|-----------|----------|
| **Type character** (Navigation mode, multi-select) | All selected cells become that value |
| **F2 / Enter** (explicit edit mode) | Only active cell edited (Excel behavior) |
| **Delete** | Clear all selected cells |
| **Backspace** | Clear all selected cells (same as Delete in Navigation) |

**Paste semantics:**

| Clipboard | Selection | Result |
|-----------|-----------|--------|
| 1×1 | Any | Fill all selected cells with value |
| N×M | 1 cell | Paste expands from active cell |
| N×M | N×M | Paste maps 1:1 |
| N×M | Larger, divisible | Tile clipboard (v2) |
| N×M | Mismatch | Error: "Paste shape mismatch" |

**Fill commands:**
- Fill Down (Ctrl+D): Copy top cell of selection to all cells below
- Fill Right (Ctrl+R): Copy left cell of selection to all cells right

### 4. Commit / Cancel Rules

| Action | In Edit Mode |
|--------|--------------|
| Enter | Commit, move down |
| Tab | Commit, move right |
| Shift+Enter | Commit, move up |
| Shift+Tab | Commit, move left |
| Escape | Cancel, restore previous value |
| Click elsewhere | Commit (default), configurable |

### 5. Undo Boundaries

One undo step per user action:

| Action | Undo Granularity |
|--------|------------------|
| Type value (replaces selection) | 1 step (all cells) |
| Paste | 1 step |
| Fill Down/Right | 1 step |
| Delete selection | 1 step |
| Single cell edit | 1 step (not per keystroke) |

**Not undoable:** Selection changes, scroll, file save

### 6. Keybinding Contexts

Bindings resolved by context priority:

1. **Palette** (highest) - when command palette open
2. **Editor** - when in edit mode
3. **Grid** - when in navigation mode
4. **Global** (lowest) - always available

Conflicts: later context wins. User bindings override defaults.

---

## Data Model

### Cell Storage

**Phase 1 (v1)**: Sparse HashMap
```rust
cells: HashMap<(usize, usize), Cell>
```
- Simple, works for mixed data
- Good for typical spreadsheet use (mostly empty)
- Memory-efficient for sparse data

**Phase 2 (v2+)**: Columnar storage for large datasets
```rust
columns: Vec<Column>
enum Column {
    Sparse(HashMap<usize, Cell>),
    DenseNumbers(Vec<f64>),
    DenseStrings(Vec<String>),
}
```
- Automatic promotion when column is >50% filled with same type
- Better cache locality for large CSV imports
- Enables SIMD operations on numeric columns

### Cell Structure

```rust
struct Cell {
    value: CellValue,
    format: Option<FormatId>,  // index into format table
    // NO per-cell style storage - use format table
}

enum CellValue {
    Empty,
    Number(f64),
    Text(String),
    Formula { source: String, ast: Expr, cached: Option<f64> },
    Error(ErrorKind),
}
```

### Sheet Limits

| Dimension | Limit | Rationale |
|-----------|-------|-----------|
| Rows | 1,048,576 (2^20) | Excel compatibility |
| Columns | 16,384 (2^14) | Excel compatibility |
| Cell text length | 32,767 chars | Excel compatibility |
| Formula length | 8,192 chars | Practical limit |
| Sheets per workbook | 256 | Practical limit |

---

## Calculation Model

### Dependency Graph

- Build DAG of cell dependencies on formula entry
- Stored separately from cell data
- Enables incremental recalc

### Recalculation Order

1. Mark dirty cell and all dependents
2. Topological sort dirty subgraph
3. Evaluate in dependency order
4. Cache results

### Volatile Functions

Functions that recalc every time:

| Function | Behavior |
|----------|----------|
| NOW() | Current datetime, recalc on any edit |
| TODAY() | Current date, recalc on any edit |
| RAND() | Random number, recalc on any edit |

Volatile cells always in dirty set. Option to disable auto-recalc for volatile-heavy sheets.

### Circular References

**Default behavior**: Error (`#CIRC!`)

**Future option**: Iterative calculation
- Max iterations: 100
- Convergence threshold: 0.001
- Disabled by default, opt-in per sheet

### Error Propagation

Errors propagate through formulas:
- `#DIV/0!` - Division by zero
- `#VALUE!` - Wrong type
- `#REF!` - Invalid reference
- `#NAME?` - Unknown function/name
- `#CIRC!` - Circular reference
- `#ERR!` - Generic error

---

## Keyboard & Command System

This is the differentiator. Get it right.

### Command Registry

All actions are commands with:
```rust
struct Command {
    id: &'static str,           // "edit.copy"
    label: &'static str,        // "Copy"
    default_binding: Option<&'static str>,  // "Ctrl+C"
    context: Context,           // When command is available
    action: fn(&mut App),       // What it does
}
```

### Keybinding Resolution

1. Check context-specific bindings first
2. Fall back to global bindings
3. Last definition wins (user overrides default)

### Contexts

```rust
enum Context {
    Global,         // Always available
    Grid,           // Navigation mode
    Editing,        // Cell edit mode
    Palette,        // Command palette open
    Dialog,         // Modal dialog open
}
```

### Binding File Format

`~/.config/excel/keybindings.json`:
```json
[
    { "key": "Ctrl+C", "command": "edit.copy", "when": "grid" },
    { "key": "Ctrl+Shift+P", "command": "palette.toggle" },
    { "key": "Ctrl+G", "command": "navigate.goto" },
    { "key": "Alt+=", "command": "formula.autosum" }
]
```

### Chord Support (v2)

Leader key sequences like vim:
- `g g` - Go to cell A1
- `g e` - Go to last cell with data
- Configurable leader key (default: none, opt-in)

---

## Plugin Architecture

### v1: Custom Functions Only

Narrow scope. One plugin type. Pure, deterministic, no IO.

```rust
// Plugin provides:
fn function_name() -> &str;
fn argument_count() -> (usize, usize);  // min, max
fn evaluate(args: &[f64]) -> Result<f64, String>;
```

**Sandboxing**: WASM modules only
- No filesystem access
- No network access
- CPU time limit per call
- Memory limit per plugin

**Distribution**: `.wasm` files in plugins directory

### Future Plugin Types (v2+)

| Type | Sandboxing | Use Case |
|------|------------|----------|
| Custom functions | WASM, pure | User-defined formulas |
| File formats | WASM, file read only | Import/export |
| Cell renderers | WASM, draw API only | Sparklines, progress bars |
| Data connectors | Native, permissioned | Database, API access |

Each type has different trust levels. Don't conflate them.

---

## File Format

### Decision: SQLite + Text Metadata

**Native format**: `.sheet` (SQLite database)

**Why SQLite**:
- Crash safety (atomic commits)
- Partial loading (query only visible range)
- Fast random access
- Incremental saves
- Single file, no temp files

**Schema**:
```sql
-- Core data
CREATE TABLE cells (
    row INTEGER,
    col INTEGER,
    value_type INTEGER,  -- 0=empty, 1=number, 2=text, 3=formula
    value_num REAL,
    value_text TEXT,
    format_id INTEGER,
    PRIMARY KEY (row, col)
);

-- Formulas stored separately for fast dependency queries
CREATE TABLE formulas (
    row INTEGER,
    col INTEGER,
    source TEXT,
    ast BLOB,  -- serialized AST
    PRIMARY KEY (row, col)
);

-- Dependencies for incremental recalc
CREATE TABLE dependencies (
    from_row INTEGER,
    from_col INTEGER,
    to_row INTEGER,
    to_col INTEGER
);

-- Metadata as JSON
CREATE TABLE meta (
    key TEXT PRIMARY KEY,
    value TEXT
);
```

**Metadata stored as JSON in `meta` table**:
- Sheet names and order
- Named ranges
- Column widths, row heights
- View state (scroll position, selection)
- Theme preference

### Diffability Solution

SQLite isn't diff-friendly. Solution:

1. **Export command**: `sheet export-text myfile.sheet > myfile.sheet.txt`
   - Outputs canonical text format (sorted cells, JSON metadata)
   - Deterministic output for same data

2. **Diff CLI**: `sheet diff old.sheet new.sheet`
   - Shows cell changes, formula changes, metadata changes
   - Machine-readable output for git integration

3. **Git integration** (optional):
   - `.gitattributes`: `*.sheet diff=sheet`
   - Custom diff driver calls `sheet diff`

### Import/Export

| Format | Import | Export | Priority |
|--------|--------|--------|----------|
| CSV | v1 | v1 | High |
| TSV | v1 | v1 | High |
| Native (.sheet) | v1 | v1 | High |
| XLSX | v2 | v2 | Medium |
| ODS | v2 | v2 | Low |
| JSON | v1 | v1 | Medium |

---

## Scripting

### Decision: Lua

**Why Lua over Python**:
- ~200KB runtime vs ~30MB Python
- <1ms startup vs ~100ms Python
- Trivially embeddable
- No dependency hell
- Designed for embedding

**Use cases**:
- Automation macros
- Custom functions (before WASM plugins)
- Batch processing
- Startup scripts

**API exposed to Lua**:
```lua
-- Cell access
sheet.get(row, col)          -- returns value
sheet.set(row, col, value)   -- sets value
sheet.get_formula(row, col)  -- returns formula string

-- Selection
sheet.selection()            -- returns {start_row, start_col, end_row, end_col}
sheet.select(r1, c1, r2, c2)

-- Commands
app.execute("edit.copy")
app.execute("navigate.goto", "A1")

-- UI
app.message("Done!")         -- status bar message
app.prompt("Enter name:")    -- input dialog
```

**Sandboxing**:
- No `os.execute`, `io.*`, `loadfile`
- Whitelist of safe standard library functions
- CPU time limit
- Memory limit

---

## Architecture

### Module Structure

```
src/
├── main.rs                 # Entry point
├── app.rs                  # Application state, message handling
├──
├── core/
│   ├── cell.rs            # Cell types and values
│   ├── sheet.rs           # Sheet data structure
│   ├── workbook.rs        # Multi-sheet container
│   └── selection.rs       # Selection model
│
├── calc/
│   ├── parser.rs          # Formula parser
│   ├── ast.rs             # Formula AST
│   ├── eval.rs            # Evaluator
│   ├── functions.rs       # Built-in functions
│   └── deps.rs            # Dependency graph
│
├── ui/
│   ├── grid.rs            # Grid rendering (canvas-based for perf)
│   ├── formula_bar.rs     # Formula bar widget
│   ├── palette.rs         # Command palette
│   ├── status.rs          # Status bar
│   └── theme.rs           # Theming system
│
├── input/
│   ├── keyboard.rs        # Keystroke routing
│   ├── commands.rs        # Command registry
│   ├── bindings.rs        # Keybinding resolution
│   └── edit.rs            # Edit mode state machine
│
├── io/
│   ├── native.rs          # SQLite format
│   ├── csv.rs             # CSV import/export
│   └── json.rs            # JSON export
│
├── scripting/
│   └── lua.rs             # Lua integration
│
└── plugins/
    └── wasm.rs            # WASM plugin host
```

### Threading Model

- **Main thread**: UI rendering, input handling
- **Calc thread**: Formula evaluation, recalc
- **IO thread**: File operations (save, load, export)

Communication via channels. UI never blocks on calc or IO.

---

## MVP Scope (v1)

### In Scope

- [x] Basic grid with cell selection
- [x] Formula bar with cell reference display
- [x] Keyboard navigation (arrows, tab, enter)
- [x] Cell editing (values and formulas)
- [x] Basic formulas (+, -, *, /, parentheses)
- [x] Cell references (A1 notation)
- [x] SUM, AVERAGE, MIN, MAX, COUNT functions
- [x] Range references (A1:B10)
- [x] Copy/cut/paste (Ctrl+C/X/V)
- [x] Command palette (Ctrl+Shift+P)
- [x] Alt+= auto-sum
- [x] Multi-cell selection (Shift+arrow, Ctrl+click)
- [x] Multi-edit (typing replaces all selected cells)
- [x] Fill Down/Right (Ctrl+D/R)
- [x] Mode state machine (Navigation/Edit/Command)
- [x] Formula syntax highlighting
- [x] Menu bar with dropdowns (File, Edit, View, Insert, Format, Data, Help)
- [x] Modern dark theme UI
- [x] Sheet tabs (visual)
- [x] CSV import/export
- [x] Configurable keybindings (file-based)
- [x] Dark/light theme toggle (functional)
- [x] Go to cell (Ctrl+G)
- [x] Undo/redo
- [x] Native format save/load (SQLite)
- [x] Context formatting bar (on selection)

### Out of Scope (v1)

- XLSX support
- Charts
- Pivot tables
- Conditional formatting
- Data validation
- Collaboration
- Printing
- Plugins
- Lua scripting

---

## Technical Decisions (Firm Stances)

| Decision | Choice | Rationale |
|----------|--------|-----------|
| UI framework | iced | Native Rust, cross-platform, GPU-accelerated |
| Native format | SQLite | Crash safety, partial loads, fast writes |
| Diff strategy | CLI tool | SQLite + export-to-text + custom diff |
| Scripting | Lua | Tiny, fast startup, embeddable |
| Plugins | WASM | Safe, portable, sandboxed |
| Formula syntax | Excel-compatible | Reduce adoption friction |
| Cell storage (v1) | Sparse HashMap | Simple, memory-efficient for typical use |
| Threading | Main + calc + IO | Never block UI |

---

## Next Steps

1. ~~**Multi-cell selection** - Shift+arrow, Shift+click, Ctrl+click~~ ✓
2. ~~**Input model refinement** - Edit mode state machine~~ ✓
3. **Undo/redo** - Command-based undo stack
4. **CSV import/export** - With encoding detection
5. **Native format** - SQLite save/load
6. **Go to cell dialog** - Ctrl+G
7. **Context formatting bar** - Appears on selection
8. **Performance baseline** - criterion benchmarks, perf HUD
9. **Keybindings file** - Load from config directory
10. **Wire up remaining menu items** - Most are placeholders

---

## Appendix: Formula Functions (v1)

| Function | Syntax | Description |
|----------|--------|-------------|
| SUM | `=SUM(range)` | Sum of values |
| AVERAGE | `=AVERAGE(range)` | Arithmetic mean |
| MIN | `=MIN(range)` | Minimum value |
| MAX | `=MAX(range)` | Maximum value |
| COUNT | `=COUNT(range)` | Count of numeric values |
| ABS | `=ABS(value)` | Absolute value |
| ROUND | `=ROUND(value, decimals)` | Round to decimals |
| IF | `=IF(cond, true, false)` | Conditional |
| AND | `=AND(a, b, ...)` | Logical AND |
| OR | `=OR(a, b, ...)` | Logical OR |
| NOT | `=NOT(value)` | Logical NOT |
| CONCAT | `=CONCAT(a, b, ...)` | Text concatenation |
| LEN | `=LEN(text)` | Text length |
| LEFT | `=LEFT(text, n)` | Left n characters |
| RIGHT | `=RIGHT(text, n)` | Right n characters |
| MID | `=MID(text, start, n)` | Substring |
| TRIM | `=TRIM(text)` | Remove whitespace |
| UPPER | `=UPPER(text)` | Uppercase |
| LOWER | `=LOWER(text)` | Lowercase |
