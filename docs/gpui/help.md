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

### Find

| Action | Shortcut |
|--------|----------|
| Open Find | Ctrl+F |
| Find next | F3 |
| Find previous | Shift+F3 |
| Close Find | Escape |

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
- Fast random access

### CSV Export

Export current sheet to CSV:
- Uses toolbar or Ctrl+Shift+S → choose .csv

---

## Modes

VisiGrid has distinct input modes:

| Mode | Behavior |
|------|----------|
| Navigation | Arrow keys move selection |
| Edit | Arrow keys move cursor in cell |
| GoTo | Type cell reference (e.g., "A1") |
| Find | Type search text |

Press Escape to return to Navigation mode.

---

## Tips

1. **Double-click** a cell to edit it
2. **Type any character** to start editing (replaces content)
3. **F2** to edit without replacing
4. **Tab** confirms edit and moves right
5. **Enter** confirms edit and moves down
6. **Escape** cancels edit
7. **Ctrl+Arrow** jumps to edge of data regions

---

## Known Limitations (gpui version)

Current limitations being addressed:

- No command palette yet (Ctrl+Shift+P)
- No Fill Down/Right (Ctrl+D/R)
- No Ctrl+Click for discontiguous selection
- No dropdown menus (toolbar only)
- No multi-sheet support yet
- No zoom
- No themes toggle
- No right-click context menu
