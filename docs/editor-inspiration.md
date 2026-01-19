# Code Editor Inspiration

Design patterns borrowed from VS Code, Sublime Text, Vim, and other modern editors. These inform VisiGrid's UX decisions.

---

## Command Palette

`Ctrl+Shift+P` to fuzzy-search every action. No menu diving.

- All commands accessible from one place
- Recent commands floated to top
- Keybinding shown inline
- Works in any mode

**Status**: Implemented (basic)

---

## Fuzzy Search Everywhere

Not just commands—search across:

- Formulas by name or description ✓ (`=` prefix)
- Named ranges ✓ (autocomplete, F12 go-to, Shift+F12 find refs, Ctrl+Shift+R rename)
- Sheet tabs (not yet - single sheet only)
- Cell contents ✓ (`@` prefix or Ctrl+F)
- Recent files ✓ (Ctrl+P or in palette)
- Settings ✓ (`#` prefix)

Command palette prefixes:
- No prefix: Commands + recent files
- `>`: Commands only
- `@`: Search cells
- `:`: Go to cell reference
- `=`: Search formula functions
- `#`: Search settings

**Status**: ✓ Complete (except sheet tabs - single sheet only)

---

## Settings as Text Files

No preferences dialog. JSON files in config directory.

```json
// keybindings.json
{ "key": "ctrl+d", "command": "selectCellBelow" }
{ "key": "ctrl+;", "command": "insertCurrentDate" }
```

```json
// settings.json
{
  "editor.fontSize": 13,
  "grid.defaultColumnWidth": 100,
  "formula.autoRecalc": true
}
```

```json
// themes/custom.json
{
  "name": "Custom",
  "colors": {
    "background": "#1a1a2e",
    "foreground": "#eaeaea"
  }
}
```

**Benefits**:
- Version controllable
- Shareable across machines
- No hidden state
- Power user friendly

**Status**: Implemented (keybindings + settings.json). Open via command palette: "Open Settings" or "Open Keyboard Shortcuts"

---

## Formula Language Server

Treat formulas like code. Apply IDE patterns.

| Feature | Description | Status |
|---------|-------------|--------|
| Autocomplete | Function names with parameter hints | ✓ Done |
| Signature help | Parameter info as you type | ✓ Done |
| Error squiggles | Red underline before execution | ✓ Done |
| Go to definition | Jump to named range definition (F12) | ✓ Done |
| Named ranges | Define names for cell ranges (Ctrl+Shift+N) | ✓ Done |
| Context help | Function docs on demand (F1) | ✓ Done |
| Find all references | Show all cells referencing a cell (Shift+F12) | ✓ Done |
| Rename symbol | Rename named range, update all references (Ctrl+Shift+R) | ✓ Done |
| Hover docs | Function signature on mouse hover | Not possible (iced limitation) |

### Implemented Features

**Autocomplete** - Type `=` and start typing a function name to see suggestions. Arrow keys to navigate, Tab/Enter to accept. Shows function syntax inline.

**Signature Help** - When typing inside function parentheses, shows:
- Function signature with all parameters
- Current parameter highlighted in bold
- Updates as you type commas between arguments

**Error Validation** - Real-time syntax checking:
- Red error message appears below formula bar
- Smart detection avoids false positives during typing
- Validates parenthesis matching, operators, function syntax
- Shows specific error messages (e.g., "Unexpected character", "Missing operand")

**Named Ranges** - Give meaningful names to cell ranges:
- `Ctrl+Shift+N` to define a name for current selection
- Names appear in formula autocomplete
- Reference by name in formulas: `=SUM(SalesData)`
- `F12` (Go to Definition) jumps to named range location when cursor is on a name in formula bar

**Context-Sensitive Help (F1)** - Press F1 while editing a formula to get detailed help:
- Shows function name, syntax, and description
- Detects the current function context automatically
- Press Escape to dismiss

**Find All References (Shift+F12)** - See all cells that reference the current cell:
- Shows list of cells with formula previews
- Click or press Enter to navigate to a reference
- Arrow keys to navigate the list

**Rename Symbol (Ctrl+Shift+R)** - Rename a named range across all formulas:
- Finds all formulas using the named range
- Shows preview of affected cells
- Updates all references automatically
- Word-boundary aware (won't rename partial matches)

### Quick Reference

| Feature | Trigger | Shortcut |
|---------|---------|----------|
| Autocomplete | Type `=` then function name | Tab/Enter to accept |
| Signature help | Type `(` after function | Auto |
| Error validation | While typing formula | Auto |
| Context help | Press F1 in formula | F1 |
| Define named range | Select cells, press shortcut | Ctrl+Shift+N |
| Go to definition | Cursor on name in formula | F12 |
| Find all references | Select cell | Shift+F12 |
| Rename symbol | Cursor on name in formula | Ctrl+Shift+R |

**Status**: ✓ All core features complete

---

## Cell Inspector Panel

Like browser DevTools for cells.

```
┌─────────────────────────────────┐
│ Inspector: D5                   │
├─────────────────────────────────┤
│ Formula:   =SUM(B2:B4)*C5       │
│ Result:    1,250.00             │
│ Type:      Number               │
├─────────────────────────────────┤
│ Precedents (depends on):        │
│   B2, B3, B4, C5                │
├─────────────────────────────────┤
│ Dependents (used by):           │
│   E5, F10                       │
├─────────────────────────────────┤
│ Format:    #,##0.00             │
│ Font:      Default, Bold        │
└─────────────────────────────────┘
```

- Toggle with `Ctrl+I` or View menu
- Click precedent/dependent to navigate
- Visual arrows on grid (optional)

**Status**: Implemented (Ctrl+Shift+I)

---

## Multi-Cursor Editing

Select non-contiguous cells, edit all simultaneously.

| Action | Keys |
|--------|------|
| Add cursor | `Ctrl+Click` |
| Add cursor above/below | `Ctrl+Alt+Up/Down` |
| Select all occurrences | `Ctrl+Shift+L` |
| Add next occurrence | `Ctrl+D` |

Typing affects all selections. Delete clears all.

**Status**: Implemented (basic multi-selection)

---

## Minimap

Bird's-eye view of large sheets.

- Shows data density (filled vs empty regions)
- Highlights formula locations
- Shows formatting regions (colors, borders)
- Click to navigate
- Highlights current viewport

**Status**: Not started (v2)

---

## Vim Mode

Optional `hjkl` navigation for vim users. Enable in settings.json:

```json
{
  "editor.vimMode": true
}
```

**Status**: ✓ Implemented (lite version)

### Available Keys (when vim mode enabled)

| Key | Action |
|-----|--------|
| `h` `j` `k` `l` | Move left/down/up/right |
| `w` | Next filled cell (right) |
| `b` | Previous filled cell (left) |
| `0` | Start of row |
| `$` | End of row (last filled cell) |
| `gg` | Top-left of sheet |
| `G` | Bottom of data |
| `{` | Previous blank row |
| `}` | Next blank row |
| `i` | Enter insert/edit mode |
| `a` | Enter insert mode (append) |
| `f` | Enter hint mode (jump) |
| `Shift+hjkl` | Extend selection |

### Behavior Notes

- Vim mode is opt-in (default off)
- When enabled, letter keys become vim commands instead of starting cell edit
- Press `i` to enter edit mode (like vim insert)
- Press `Escape` to return to navigation (normal mode)
- Visual mode not yet implemented (use Shift+hjkl for selection)

---

## Plugin Architecture

Extend without forking.

| Plugin Type | Sandboxing | Example |
|-------------|------------|---------|
| Custom functions | WASM, pure | `=MYCOMPANY.RATE(x)` |
| Data connectors | Native, permissioned | PostgreSQL, REST API |
| Cell renderers | WASM, draw API | Sparklines, progress bars |
| Export formats | WASM, file write | Custom CSV dialect |
| Themes | JSON, static | Color schemes |

**Status**: Not started (v2)

---

## Workspaces

Remember state per project.

Saved state:
- Open files and tabs
- Scroll position per sheet
- Selection state
- Zoom level
- Panel layout
- Recent commands

```
~/.config/visigrid/workspaces/
  <hash>.json  # One per project directory
```

**How it works**:
- Auto-detects project by looking for `.visigrid` marker file up the directory tree
- Falls back to current working directory if no marker found
- Loads workspace session first, then falls back to global session
- To create a project: `touch /path/to/project/.visigrid`

**Status**: ✓ Implemented

---

## Diffable Format

Git-friendly native format.

**Options**:

1. **SQLite + export tool** (current choice)
   - Fast, crash-safe
   - `visigrid diff old.sheet new.sheet` for text diff

2. **Structured text** (alternative)
   - Line-based format for easy diffs
   - Slower for large files

### CLI Diff Tool

```bash
visigrid diff old.sheet new.sheet
```

Output format (unified diff style):
```
--- old.sheet
+++ new.sheet
@@ A1 @@
-100
+150
@@ B3 @@
+=SUM(A1:A10)
```

Shows added (`+`), removed (`-`), and changed cells with their cell references.

**Status**: ✓ Implemented (SQLite format + CLI diff tool)

---

## Integrated Scripting

Visible REPL console, not hidden macros. Uses Lua for scripting.

```
┌─────────────────────────────────┐
│ > sheet:get_a1("A1")            │
│ 42                              │
│ > for r = 1, 100 do             │
│ ...   sheet:set_value(r, 1, r*2)│
│ ... end                         │
│ nil                             │
│ 100 cells modified              │
│ >                               │
└─────────────────────────────────┘
```

- `Ctrl+Shift+L` to toggle REPL panel
- Command history (up/down arrows)
- Sandboxed execution (no file/OS access)
- Instruction limits and wall-clock timeout
- Typed values (numbers, strings, booleans, errors)

### Lua API

| Method | Description |
|--------|-------------|
| `sheet:get_value(row, col)` | Get typed value (1-indexed) |
| `sheet:get_display(row, col)` | Get formatted display string |
| `sheet:get_formula(row, col)` | Get formula or nil |
| `sheet:set_value(row, col, val)` | Set cell value |
| `sheet:set_formula(row, col, formula)` | Set formula |
| `sheet:get_a1("A1")` | Get value using A1 notation |
| `sheet:set_a1("A1", val)` | Set value using A1 notation |
| `sheet:clear(row, col)` | Clear cell |
| `sheet:rows()` / `sheet:cols()` | Sheet dimensions |
| `sheet:range("A1:C10")` | Get range object |

**Status**: ✓ Implemented (Ctrl+Shift+L)

---

## Additional Ideas (Not in Original)

### Problems Panel

Aggregate view of all formula errors.

```
Problems (3)
  D5: #REF! - Invalid cell reference
  E10: #DIV/0! - Division by zero
  F2: #NAME? - Unknown function VLOKUP
```

Click to navigate. Auto-updates on edit.

**Status**: Implemented (Ctrl+Shift+M)

### Breadcrumbs

Context bar showing current location.

```
[Workbook.sheet] › [D5] › [=SUM(B2:B4)]
```

- File segment → clickable, opens file dialog
- Cell reference → clickable, opens Go To dialog
- Value/formula preview (truncated to 40 chars)

**Status**: ✓ Implemented

### Zen Mode

Distraction-free editing.

- `F11` to toggle (or via command palette)
- `Escape` to exit
- Hides all panels (menu, formula bar, format bar, sheet tabs)
- Full-screen grid view

**Status**: ✓ Implemented

### Session Persistence

Auto-save state on quit, restore on launch.

Saved state:
- Current file
- Scroll position
- Active cell/selection
- Split view configuration
- UI settings (dark mode, zen mode, panels)

```bash
visigrid                    # Restore previous session
visigrid --no-restore       # Start fresh
visigrid -n                 # Start fresh (short form)
visigrid file.sheet         # Open specific file (skip session)
```

Session stored at: `~/.config/visigrid/session.json`

**Status**: ✓ Implemented

### Paste Special

Excel-like paste options (`Ctrl+Shift+V`):

**Paste Types**:
- All (default)
- Values only (no formulas)
- Formulas only
- Formats only

**Operations**:
- None (just paste)
- Add / Subtract / Multiply / Divide (apply to existing values)

**Options**:
- Transpose (swap rows/columns)
- Skip blanks

**Keyboard shortcuts in dialog**:
- `↑/↓` navigate, `Tab` switch sections
- `T` toggle transpose, `B` toggle skip blanks
- `A/V/F/O` quick select paste type
- `+/-/*/` quick select operation
- `Enter` paste, `Esc` cancel

**Status**: ✓ Implemented

### Split View

Edit two sheets side by side.

- Vertical or horizontal split
- Independent scroll
- Link selection (optional)
- Compare mode

**Status**: Implemented (basic - Ctrl+\ toggle, Ctrl+W switch pane)

### Quick Open

`Ctrl+P` to open recent files.

- Fuzzy search file names
- Shows recent files first
- Preview on hover

**Status**: Implemented

---

## Implementation Priority (by Impact)

Sorted by how much each feature improves daily productivity.

| Rank | Feature | Why It Matters | Status |
|------|---------|----------------|--------|
| 1 | **Formula Language Server** | Transforms formula writing from "guess and pray" to guided editing. Autocomplete alone saves massive time. Error squiggles catch mistakes before they propagate. | ✓ Complete |
| 2 | **Cell Inspector** | "Why did this change?" is the #1 debugging question. Seeing precedents/dependents instantly is like having a debugger for your data. | ✓ Done |
| 3 | **Command Palette** | Foundation for everything. Makes features discoverable. Teaches shortcuts. | ✓ Done |
| 4 | **Multi-selection + Multi-edit** | Core differentiator. Edit 50 cells at once instead of copy-paste loops. | ✓ Done |
| 5 | **Problems Panel** | See all errors in one place instead of hunting. Click to fix. Essential for large sheets. | ✓ Done |
| 6 | **Fuzzy Search Everywhere** | Find anything instantly. Cells, formulas, named ranges, settings. Currently you're blind in large workbooks. | ✓ Complete |
| 7 | **Settings as Files** | Power users expect it. Teams can share configs. Version control friendly. | ✓ Done |
| 8 | **Quick Open (Ctrl+P)** | Fast file switching. Table stakes for any productivity tool. | ✓ Done |
| 9 | **Integrated Scripting** | Automate repetitive tasks. Visible REPL beats hidden macros. | ✓ Done |
| 10 | **Split View** | Compare sheets, reference while editing. Common workflow, painful without it. | ✓ Done (basic) |
| 11 | **Workspaces** | Project context. Nice for heavy users, not critical for most. | ✓ Done |
| 12 | **Minimap** | Orientation in large sheets. Helpful but you can live without it. | Not started |
| 13 | **Vim Mode** | Niche audience, but those users are vocal and loyal. Low effort if designed well. | ✓ Done (lite) |
| 14 | **Zen Mode** | Polish feature. Nice to have, zero impact on core workflow. | ✓ Done |
| 15 | **Breadcrumbs** | Context bar for navigation. Shows file, cell, and value. | ✓ Done |
| 16 | **Session Persistence** | Auto-save/restore on quit/launch. | ✓ Done |
| 17 | **Paste Special** | Excel-like paste options (values, formulas, transpose, operations). | ✓ Done |

### The Big Three

If I had to pick three features that would most transform the experience:

1. ~~**Formula Language Server**~~ ✓ — Makes formulas feel like code, not magic strings
2. ~~**Cell Inspector**~~ ✓ — Finally understand your spreadsheet's structure
3. ~~**Problems Panel**~~ ✓ — Stop playing whack-a-mole with errors

**All three core features are now complete!** VisiGrid has evolved from "fast Excel clone" into "IDE for tabular data."

---

## Principles

1. **Keyboard first** - Every action reachable without mouse
2. **Discoverable** - Command palette teaches shortcuts
3. **Configurable** - Power users can customize everything
4. **Transparent** - Settings in text files, not hidden databases
5. **Fast** - Latency is a feature
