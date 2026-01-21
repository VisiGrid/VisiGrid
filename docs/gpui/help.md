# VisiGrid Help (gpui version)

Quick reference for current functionality.

---

## Getting Started

### Navigation

| Action | Shortcut |
|--------|----------|
| Move between cells | Arrow keys |
| Move to next cell | Tab |
| Move to previous cell | Shift+Tab |
| Jump to edge of data | Ctrl+Arrow |
| Go to specific cell | Ctrl+G |
| Go to cell A1 | Ctrl+Home |
| Go to last cell | Ctrl+End |
| Page up/down | Page Up/Down |

### Editing

| Action | Shortcut |
|--------|----------|
| Start editing | F2 or type any character |
| Confirm edit (move down) | Enter |
| Confirm edit (move right) | Tab |
| Cancel edit | Escape |
| Delete cell contents | Delete |

### Selection

| Action | Shortcut |
|--------|----------|
| Extend selection | Shift+Arrow |
| Select all | Ctrl+A |
| Extend by clicking | Shift+Click |
| Add to selection | Ctrl+Click |

### Multi-Edit

Select multiple cells, then type to edit all at once:

| Action | Shortcut |
|--------|----------|
| Apply edit to all cells | Enter |
| Fill from primary cell | Ctrl+Enter |
| Cancel | Escape |

**Live Preview:** All selected cells show what they'll receive while you type.

**Formula Shifting:** Relative references adjust automatically (e.g., `=A1*2` becomes `=B2*2` in the next cell).

### Clipboard

| Action | Shortcut |
|--------|----------|
| Copy | Ctrl+C |
| Cut | Ctrl+X |
| Paste | Ctrl+V |

### Undo/Redo

| Action | Shortcut |
|--------|----------|
| Undo | Ctrl+Z |
| Redo | Ctrl+Y or Ctrl+Shift+Z |

### Formatting

| Action | Shortcut |
|--------|----------|
| Bold | Ctrl+B |
| Italic | Ctrl+I |
| Underline | Ctrl+U |

### File Operations

| Action | Shortcut |
|--------|----------|
| New file | Ctrl+N |
| Open file | Ctrl+O |
| Save | Ctrl+S |
| Save As | Ctrl+Shift+S |
| Export CSV | File menu or Palette |
| Export TSV | File menu or Palette |
| Export JSON | File menu or Palette |

### Find

| Action | Shortcut |
|--------|----------|
| Open Find | Ctrl+F |
| Find next | F3 |
| Find previous | Shift+F3 |
| Close Find | Escape |

### View

| Action | Shortcut |
|--------|----------|
| Zen Mode (distraction-free) | F11 |
| Command Palette | Ctrl+Shift+P |
| Fuzzy Finder | Ctrl+P |

---

## Toolbar Buttons

The toolbar provides quick access to common actions:

| Button | Action |
|--------|--------|
| New | Create new spreadsheet |
| Open | Open existing file |
| Save | Save current file |
| Undo | Undo last action |
| Redo | Redo undone action |
| Cut | Cut selection |
| Copy | Copy selection |
| Paste | Paste clipboard |
| B | Toggle bold |
| I | Toggle italic |
| U | Toggle underline |
| Find | Open find dialog |
| GoTo | Open go-to dialog |

---

## Formulas

### Basic Syntax

Formulas start with `=`:

```
=A1+B1           Simple addition
=SUM(A1:A10)     Sum a range
=IF(A1>0,"Yes","No")   Conditional
```

### Cell References

| Type | Example | Behavior when copied |
|------|---------|---------------------|
| Relative | A1 | Adjusts |
| Absolute | $A$1 | Stays fixed |
| Mixed | $A1 or A$1 | Partial adjustment |

### Common Functions

| Function | Example | Description |
|----------|---------|-------------|
| SUM | =SUM(A1:A10) | Add numbers |
| AVERAGE | =AVERAGE(A1:A10) | Mean value |
| COUNT | =COUNT(A1:A10) | Count numbers |
| IF | =IF(A1>0,1,0) | Conditional |
| VLOOKUP | =VLOOKUP(A1,B:C,2,0) | Lookup value |
| TODAY | =TODAY() | Current date |

### Operators

| Operator | Meaning |
|----------|---------|
| + | Add |
| - | Subtract |
| * | Multiply |
| / | Divide |
| & | Concatenate text |
| < > = | Comparison |

---

## File Formats

### Native Format (.sheet)

SQLite-based format with:
- Cell values and formulas
- Formatting
- Sheet metadata
- Named ranges
- Column widths
- Fast random access

### Excel Import (xlsx/xls/xlsb/ods)

VisiGrid can open Excel files:
- One-way import (Save As .sheet to keep changes)
- Background import for large files (UI stays responsive)
- Import Report shows fidelity info
- Unsupported functions are tracked and reported
- Dates, formulas, and formatting are preserved

### CSV/TSV/JSON Export

- Open CSV and TSV files directly
- Export via File menu or Command Palette:
  - **CSV** — comma-separated values
  - **TSV** — tab-separated values
  - **JSON** — array of arrays format

---

## Modes

VisiGrid has distinct input modes:

| Mode | Behavior |
|------|----------|
| Navigation | Arrow keys move selection |
| Edit | Arrow keys move cursor in cell |
| GoTo | Type cell reference (e.g., "A1") |
| Find | Type search text |
| Hint | Type letters to jump to labeled cells |

Press Escape to return to Navigation mode.

---

## Optional Navigation Modes

Enable these in Preferences (Ctrl+,) → Navigation. Both are off by default to preserve the "type to start editing" behavior.

### Keyboard Hints (Vimium-style)

Press `g` to show letter labels on visible cells, then type the letters to jump directly.

| Shortcut | Action |
|----------|--------|
| g | Enter hint mode (show labels) |
| gg | Go to cell A1 |
| a-z, aa-zz | Jump to labeled cell |
| Escape | Cancel hint mode |
| Backspace | Delete last character |

The status bar shows `HINT` when active, plus the letters you've typed.

### Vim Mode

Navigate without entering edit mode accidentally.

| Shortcut | Action |
|----------|--------|
| h / j / k / l | Move left/down/up/right |
| i | Enter edit mode |
| 0 | Jump to column A |
| $ | Jump to last data column in row |
| w | Jump right (next data edge) |
| b | Jump left (prev data edge) |

The status bar shows `VIM` instead of `NAV` when enabled.

---

## Named Ranges

Create named ranges to make formulas more readable:

| Action | Shortcut |
|--------|----------|
| Create named range | Ctrl+Shift+N |
| Rename symbol | Ctrl+Shift+R |
| Extract to named range | Select range, then Ctrl+Shift+N |

Named ranges appear in:
- Formula autocomplete
- Fuzzy finder (Ctrl+P)
- Inspector panel

---

## Multi-Sheet

| Action | Shortcut |
|--------|----------|
| Next sheet | Ctrl+PageDown |
| Previous sheet | Ctrl+PageUp |
| New sheet | Shift+F11 |
| Sheet menu | Right-click tab |

---

## Themes & Preferences

| Action | Shortcut |
|--------|----------|
| Theme picker | Ctrl+K Ctrl+T |
| Preferences | Ctrl+, |
| Open Keybindings | Command Palette |

---

## Zen Mode

Press **F11** to toggle Zen Mode — a distraction-free view that hides the menu bar, formula bar, and status bar. Only the grid remains.

| Action | Shortcut |
|--------|----------|
| Toggle Zen Mode | F11 |
| Exit Zen Mode | Escape or F11 |

Also available via Command Palette: "Toggle Zen Mode"

---

## Link Detection

VisiGrid detects URLs, email addresses, and file paths in cells and lets you open them directly.

| Type | Example | Opens with |
|------|---------|------------|
| URL | `https://example.com` | Default browser |
| Email | `user@example.com` | Default mail client (mailto:) |
| File path | `/home/user/doc.pdf` | System default handler |

**Usage:** Select a cell containing a link, then press **Ctrl+Enter** to open it.

**Detection rules (conservative):**
- URLs must have a scheme (http://, https://, ftp://, file://)
- Emails must have user@domain.tld format (dot required in domain)
- File paths must start with `/` or `~` and the file must exist

The status bar shows a hint like "URL: Ctrl+Enter to open" when a link is detected.

---

## Transform Commands

Available via Command Palette (Ctrl+Shift+P):

| Command | Description |
|---------|-------------|
| Transform: Trim Whitespace | Remove leading/trailing spaces from selected cells |
| Select: Blanks in Region | Select all empty cells within current selection |

---

## Keybinding Customization

Customize shortcuts via JSON file:

**Location:** `~/.config/visigrid/keybindings.json`

**Access:** Command Palette → "Open Keybindings (JSON)"

**Example:**
```json
{
  "ctrl-shift-d": "edit.filldown",
  "ctrl-;": "edit.trim"
}
```

User keybindings override defaults. Restart required after changes.

---

## Inspector Panel

Toggle with Ctrl+Shift+I. Shows:
- Cell address and value
- Formula (if present)
- Format information
- Named range usage

---

## Tips

1. **Double-click** a cell to edit it
2. **Type any character** to start editing (replaces content)
3. **F2** to edit without replacing
4. **Tab** confirms edit and moves right
5. **Enter** confirms edit and moves down
6. **Escape** cancels edit
7. **Ctrl+Arrow** jumps to edge of data regions
8. **Ctrl+P** opens fuzzy finder for cells and named ranges
9. **Ctrl+Shift+P** opens command palette

---

## Known Limitations

Current limitations:

- No cross-sheet references (=Sheet2!A1)
- No zoom (Ctrl++/-)
- No freeze panes
- No XLSX export (import only, use CSV/TSV/JSON for export)
- No Replace dialog (Ctrl+H)
- No Format Cells dialog (Ctrl+1)
