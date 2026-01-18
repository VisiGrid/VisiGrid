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
- Named ranges (not yet)
- Sheet tabs (not yet - single sheet only)
- Cell contents ✓ (`@` prefix or Ctrl+F)
- Recent files ✓ (Ctrl+P or in palette)
- Settings (not yet)

Command palette prefixes:
- No prefix: Commands + recent files
- `>`: Commands only
- `@`: Search cells
- `:`: Go to cell reference
- `=`: Search formula functions

**Status**: Mostly implemented

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

| Feature | Description |
|---------|-------------|
| Autocomplete | Function names with parameter hints |
| Hover docs | Function signature + description on hover |
| Error squiggles | Red underline before execution |
| Go to definition | Jump to named range definition |
| Find all references | Show all cells referencing a cell |
| Rename symbol | Rename named range, update all references |
| Signature help | Parameter info as you type |

**Status**: In progress - Autocomplete implemented (type function name after `=`)

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

Optional `hjkl` navigation for vim users.

| Mode | Behavior |
|------|----------|
| Normal | Navigation, commands |
| Insert | Cell editing |
| Visual | Selection |
| Command | `:` commands |

Motions:
- `w` / `b` - Next/prev filled cell
- `gg` / `G` - Top/bottom of data
- `0` / `$` - Start/end of row
- `{` / `}` - Prev/next blank row

**Status**: Not started (v2, optional)

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
  project-a.workspace.json
  project-b.workspace.json
```

**Status**: Not started (v2)

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

**Status**: Implemented (SQLite), CLI diff tool planned

---

## Integrated Scripting

Visible REPL console, not hidden macros.

```
┌─────────────────────────────────┐
│ > sheet.get("A1")               │
│ 42                              │
│ > for r in 1..100:              │
│ ...   sheet.set(r, 1, r * 2)    │
│ Done: 100 cells modified        │
│ >                               │
└─────────────────────────────────┘
```

- `Ctrl+`` to toggle console
- History, autocomplete
- Script files in project directory
- Runs in sandbox (no arbitrary file access)

**Status**: Not started (v2)

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
[Workbook.sheet] > [Sheet1] > [D5] > [=SUM(B2:B4)]
```

Clickable for navigation.

### Zen Mode

Distraction-free editing.

- `F11` to toggle (or via command palette)
- `Escape` to exit
- Hides all panels (menu, formula bar, format bar, sheet tabs)
- Full-screen grid view

**Status**: Implemented

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
| 1 | **Formula Language Server** | Transforms formula writing from "guess and pray" to guided editing. Autocomplete alone saves massive time. Error squiggles catch mistakes before they propagate. | In progress (autocomplete done) |
| 2 | **Cell Inspector** | "Why did this change?" is the #1 debugging question. Seeing precedents/dependents instantly is like having a debugger for your data. | ✓ Done |
| 3 | **Command Palette** | Foundation for everything. Makes features discoverable. Teaches shortcuts. | ✓ Done |
| 4 | **Multi-selection + Multi-edit** | Core differentiator. Edit 50 cells at once instead of copy-paste loops. | ✓ Done |
| 5 | **Problems Panel** | See all errors in one place instead of hunting. Click to fix. Essential for large sheets. | ✓ Done |
| 6 | **Fuzzy Search Everywhere** | Find anything instantly. Cells, formulas, named ranges, sheets. Currently you're blind in large workbooks. | Mostly done |
| 7 | **Settings as Files** | Power users expect it. Teams can share configs. Version control friendly. | ✓ Done |
| 8 | **Quick Open (Ctrl+P)** | Fast file switching. Table stakes for any productivity tool. | ✓ Done |
| 9 | **Integrated Scripting** | Automate repetitive tasks. Visible REPL beats hidden macros. | Not started |
| 10 | **Split View** | Compare sheets, reference while editing. Common workflow, painful without it. | ✓ Done (basic) |
| 11 | **Workspaces** | Project context. Nice for heavy users, not critical for most. | Not started |
| 12 | **Minimap** | Orientation in large sheets. Helpful but you can live without it. | Not started |
| 13 | **Vim Mode** | Niche audience, but those users are vocal and loyal. Low effort if designed well. | Not started |
| 14 | **Zen Mode** | Polish feature. Nice to have, zero impact on core workflow. | ✓ Done |

### The Big Three

If I had to pick three features that would most transform the experience:

1. **Formula Language Server** (in progress) — Makes formulas feel like code, not magic strings
2. ~~**Cell Inspector**~~ ✓ — Finally understand your spreadsheet's structure
3. ~~**Problems Panel**~~ ✓ — Stop playing whack-a-mole with errors

These three turn VisiGrid from "fast Excel clone" into "IDE for tabular data."

---

## Principles

1. **Keyboard first** - Every action reachable without mouse
2. **Discoverable** - Command palette teaches shortcuts
3. **Configurable** - Power users can customize everything
4. **Transparent** - Settings in text files, not hidden databases
5. **Fast** - Latency is a feature
