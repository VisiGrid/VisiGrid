# VisiGrid Help

VisiGrid is a modern spreadsheet application that combines the familiarity of Excel with IDE-like features from code editors like VS Code.

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [Navigation](#navigation)
3. [Vim Mode](#vim-mode)
4. [Selection](#selection)
5. [Editing Cells](#editing-cells)
6. [Paste Special](#paste-special)
7. [Formulas](#formulas)
8. [Formula Language Server](#formula-language-server)
9. [Named Ranges](#named-ranges)
10. [Lua Scripting](#lua-scripting)
11. [Command Palette](#command-palette)
12. [Find and Search](#find-and-search)
13. [Cell Inspector](#cell-inspector)
14. [Problems Panel](#problems-panel)
15. [View Modes](#view-modes)
16. [Formatting](#formatting)
17. [File Operations](#file-operations)
18. [CLI Tools](#cli-tools)
19. [Settings and Customization](#settings-and-customization)
20. [Quick Reference](#quick-reference)

---

## Getting Started

### Opening VisiGrid

Run `visigrid` from the command line. You can optionally pass a file path to open:

```bash
visigrid                    # Open with empty sheet
visigrid myfile.vgrid       # Open existing file
```

### The Interface

```
┌─────────────────────────────────────────────────────────────┐
│ File  Edit  View  Format  Help                    [Menu]    │
├─────────────────────────────────────────────────────────────┤
│ A1 ▼ │ =SUM(B2:B10)                          [Formula Bar]  │
├─────────────────────────────────────────────────────────────┤
│ B I U │ Left Center Right │ $ % │ +.00 -.00   [Format Bar]  │
├───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┤
│   │ A │ B │ C │ D │ E │ F │ G │ H │ I │ J │ K │ L │ M │   │
├───┼───┼───┼───┼───┼───┼───┼───┼───┼───┼───┼───┼───┼───┼───┤
│ 1 │   │   │   │   │   │   │   │   │   │   │   │   │   │   │
│ 2 │   │   │   │   │   │   │   │   │   │   │   │   │   │   │
│...│   │   │   │   │   │   │   │   │   │   │   │   │   │   │
└───┴───┴───┴───┴───┴───┴───┴───┴───┴───┴───┴───┴───┴───┴───┘
```

---

## Navigation

### Moving Around the Grid

| Action | Shortcut |
|--------|----------|
| Move one cell | Arrow keys |
| Move to next cell | `Tab` |
| Move to previous cell | `Shift+Tab` |
| Move down (confirm entry) | `Enter` |
| Move up | `Shift+Enter` |
| Jump to edge of data | `Ctrl+Arrow` |
| Go to first cell (A1) | `Ctrl+Home` |
| Go to last used cell | `Ctrl+End` |
| Page up/down | `Page Up` / `Page Down` |

### Go To Cell

Press `Ctrl+G` to open the Go To dialog. Type a cell reference like `A1`, `B25`, or `AA100` and press Enter.

You can also use the command palette (`:` prefix) - type `:A1` to jump directly.

### Keyboard Hints (Vimium-style)

Press `g` to enter hint mode. Letter hints appear on all visible cells:

```
┌───────┬───────┬───────┬───────┐
│   A   │   B   │   C   │   D   │
├───────┼───────┼───────┼───────┤
│   E   │   F   │   G   │   H   │
└───────┴───────┴───────┴───────┘
```

Type the hint letters to jump instantly to that cell. For large grids, hints use two-letter combinations (aa, ab, etc.). Press `Escape` to cancel.

### Opening Links

`Ctrl+Click` on a cell containing a URL, email address, or file path to open it:
- URLs: `https://example.com` opens in browser
- Emails: `user@example.com` opens mail client
- Paths: `/home/user/file.txt` or `~/Documents/file.pdf` opens in default app

---

## Vim Mode

VisiGrid supports optional Vim-style navigation for power users. Enable it in settings:

```json
{
  "editor.vimMode": true
}
```

### Vim Keys (when enabled)

| Key | Action |
|-----|--------|
| `h` `j` `k` `l` | Move left/down/up/right |
| `Shift+hjkl` | Extend selection |
| `w` | Next filled cell (right) |
| `b` | Previous filled cell (left) |
| `0` | Start of row |
| `$` | End of row (last filled cell) |
| `gg` | Top-left of sheet |
| `G` | Bottom of data |
| `{` | Previous blank row |
| `}` | Next blank row |
| `i` | Enter insert/edit mode |
| `a` | Append (edit with cursor at end) |
| `f` | Enter hint mode (jump to cell) |
| `Escape` | Return to normal mode |

**Note:** When Vim mode is enabled, pressing letter keys in navigation mode executes Vim commands instead of starting cell edit. Press `i` to start typing.

---

## Selection

### Basic Selection

| Action | Shortcut |
|--------|----------|
| Extend selection | `Shift+Arrow` |
| Extend to edge of data | `Ctrl+Shift+Arrow` |
| Select all cells | `Ctrl+A` |
| Select entire column | `Ctrl+Space` |
| Select entire row | `Shift+Space` |
| Select to end of data | `Ctrl+Shift+End` |

### Multi-Selection

| Action | Method |
|--------|--------|
| Add cell to selection | `Ctrl+Click` |
| Extend selection | `Shift+Click` |
| Select column | Click column header |
| Select row | Click row number |

### Multi-Edit (Edit Multiple Cells at Once)

When multiple cells are selected, typing affects all selected cells simultaneously with **live preview** and **smart formula shifting**.

**How it works:**
1. Select multiple cells using `Ctrl+Click` or `Shift+Click`
2. Start typing - all selected cells show a preview of what they'll receive
3. Press `Enter` to apply to all cells, or `Escape` to cancel

**Formula Reference Shifting:**

When entering a formula with multi-selection, relative references automatically shift based on each cell's position relative to the primary cell:

```
Example: Select D1, E2, F3 (Ctrl+Click), type =A1*2

While typing, you see:
  D1: =A1*2     (active edit cell)
  E2: =B2*2     (dimmed preview - shifted by +1 row, +1 col)
  F3: =C3*2     (dimmed preview - shifted by +2 rows, +2 cols)

Press Enter → all three formulas are applied with shifted references
```

**Reference types:**
- **Relative** (`A1`): Shifts with each cell's position
- **Absolute** (`$A$1`): Stays fixed in all cells
- **Mixed** (`$A1` or `A$1`): Partially shifts

**Fill Selection (Ctrl+Enter):**

In navigation mode with multi-selection, press `Ctrl+Enter` to fill all selected cells with the primary cell's content (with formula shifting).

| Action | Shortcut |
|--------|----------|
| Apply edit to all selected cells | `Enter` |
| Fill selection from primary cell | `Ctrl+Enter` |
| Cancel multi-edit | `Escape` |

**Status bar hints:**
- Navigation mode: "Type to edit all · Ctrl+Enter to fill"
- Edit mode: "Enter to apply · Esc to cancel"

---

## Editing Cells

### Entering Data

Simply start typing to enter data in the selected cell. The previous content will be replaced.

| Action | Shortcut |
|--------|----------|
| Edit existing cell | `F2` |
| Confirm and move down | `Enter` |
| Confirm and move right | `Tab` |
| Confirm and move up | `Shift+Enter` |
| Confirm and move left | `Shift+Tab` |
| Cancel editing | `Escape` |
| Clear cell contents | `Delete` |

### Clipboard Operations

| Action | Shortcut |
|--------|----------|
| Copy | `Ctrl+C` |
| Cut | `Ctrl+X` |
| Paste | `Ctrl+V` |
| Undo | `Ctrl+Z` |
| Redo | `Ctrl+Y` |

### Fill Operations

| Action | Shortcut |
|--------|----------|
| Fill down | `Ctrl+D` |
| Fill right | `Ctrl+R` |

Select a range where the first cell has content, then use Fill Down/Right to copy that content to all selected cells.

### Insert and Delete

| Action | Shortcut |
|--------|----------|
| Insert row/column | `Ctrl+Shift+=` |
| Delete row/column | `Ctrl+-` |

When a full row is selected (via `Shift+Space`), these commands affect rows. When a full column is selected (via `Ctrl+Space`), they affect columns.

### Date and Time

| Action | Shortcut |
|--------|----------|
| Insert current date | `Ctrl+;` |
| Insert current time | `Ctrl+Shift+;` |

---

## Paste Special

Press `Ctrl+Alt+V` after copying cells to open the Paste Special dialog.

### Paste Types

| Type | Description |
|------|-------------|
| All | Paste everything (default) |
| Values | Paste computed values only (no formulas) |
| Formulas | Paste formulas only (no formatting) |
| Formats | Paste formatting only (no data) |

### Operations

Apply a mathematical operation to existing cell values:

| Operation | Effect |
|-----------|--------|
| None | Replace existing values |
| Add | Add pasted values to existing |
| Subtract | Subtract pasted values from existing |
| Multiply | Multiply existing by pasted values |
| Divide | Divide existing by pasted values |

### Options

| Option | Description |
|--------|-------------|
| Transpose | Swap rows and columns |
| Skip Blanks | Don't overwrite cells with blank source cells |

### Keyboard Shortcuts in Dialog

| Key | Action |
|-----|--------|
| `↑/↓` | Navigate options |
| `Tab` | Switch between sections |
| `A/V/F/O` | Quick select paste type |
| `+/-/*/` | Quick select operation |
| `T` | Toggle transpose |
| `B` | Toggle skip blanks |
| `Enter` | Execute paste |
| `Escape` | Cancel |

---

## Formulas

### Entering Formulas

Start a formula with `=`. For example:
- `=A1+B1` - Add two cells
- `=SUM(A1:A10)` - Sum a range
- `=AVERAGE(B1:B5)*2` - Combine functions and operators

### Available Functions

| Function | Syntax | Description |
|----------|--------|-------------|
**Math Functions:**
| Function | Syntax | Description |
|----------|--------|-------------|
| SUM | `=SUM(range)` | Adds all numbers in a range |
| AVERAGE | `=AVERAGE(range)` | Returns the arithmetic mean |
| MIN | `=MIN(range)` | Returns the smallest value |
| MAX | `=MAX(range)` | Returns the largest value |
| COUNT | `=COUNT(range)` | Counts numbers in a range |
| COUNTA | `=COUNTA(range)` | Counts non-empty cells |
| ABS | `=ABS(value)` | Returns the absolute value |
| ROUND | `=ROUND(value, decimals)` | Rounds to specified decimals |
| INT | `=INT(value)` | Truncates to integer |
| MOD | `=MOD(number, divisor)` | Returns remainder |
| POWER | `=POWER(base, exp)` | Returns base^exp |
| SQRT | `=SQRT(value)` | Square root |
| PRODUCT | `=PRODUCT(range)` | Multiplies all values |
| MEDIAN | `=MEDIAN(range)` | Returns median value |

**Logical Functions:**
| Function | Syntax | Description |
|----------|--------|-------------|
| IF | `=IF(cond, true, false)` | Conditional result |
| AND | `=AND(cond1, cond2, ...)` | TRUE if all true |
| OR | `=OR(cond1, cond2, ...)` | TRUE if any true |
| NOT | `=NOT(condition)` | Reverses boolean |
| IFERROR | `=IFERROR(val, err_val)` | Error handling |
| ISBLANK | `=ISBLANK(cell)` | TRUE if empty |
| ISNUMBER | `=ISNUMBER(value)` | TRUE if number |
| ISTEXT | `=ISTEXT(value)` | TRUE if text |

**Text Functions:**
| Function | Syntax | Description |
|----------|--------|-------------|
| CONCATENATE | `=CONCATENATE(a, b, ...)` | Joins text |
| LEFT | `=LEFT(text, n)` | Left n characters |
| RIGHT | `=RIGHT(text, n)` | Right n characters |
| MID | `=MID(text, start, n)` | Middle characters |
| LEN | `=LEN(text)` | Length of text |
| UPPER | `=UPPER(text)` | Uppercase |
| LOWER | `=LOWER(text)` | Lowercase |
| TRIM | `=TRIM(text)` | Remove extra spaces |
| FIND | `=FIND(find, in)` | Find position |
| SUBSTITUTE | `=SUBSTITUTE(text, old, new)` | Replace text |

**Conditional Functions:**
| Function | Syntax | Description |
|----------|--------|-------------|
| SUMIF | `=SUMIF(range, criteria, sum_range)` | Conditional sum |
| COUNTIF | `=COUNTIF(range, criteria)` | Conditional count |
| COUNTBLANK | `=COUNTBLANK(range)` | Count empty cells |

**Lookup & Reference Functions:**
| Function | Syntax | Description |
|----------|--------|-------------|
| VLOOKUP | `=VLOOKUP(value, range, col, [sorted])` | Vertical lookup |
| HLOOKUP | `=HLOOKUP(value, range, row, [sorted])` | Horizontal lookup |
| INDEX | `=INDEX(range, row, [col])` | Value at position |
| MATCH | `=MATCH(value, range, [type])` | Find position |
| ROW | `=ROW([cell])` | Row number |
| COLUMN | `=COLUMN([cell])` | Column number |
| ROWS | `=ROWS(range)` | Count rows |
| COLUMNS | `=COLUMNS(range)` | Count columns |

**Date & Time Functions:**
| Function | Syntax | Description |
|----------|--------|-------------|
| TODAY | `=TODAY()` | Current date |
| NOW | `=NOW()` | Current date and time |
| DATE | `=DATE(year, month, day)` | Create date |
| YEAR | `=YEAR(date)` | Extract year |
| MONTH | `=MONTH(date)` | Extract month (1-12) |
| DAY | `=DAY(date)` | Extract day |
| WEEKDAY | `=WEEKDAY(date, [type])` | Day of week (1-7) |
| DATEDIF | `=DATEDIF(start, end, unit)` | Date difference |
| EDATE | `=EDATE(date, months)` | Add months |
| EOMONTH | `=EOMONTH(date, months)` | End of month |
| HOUR | `=HOUR(time)` | Extract hour |
| MINUTE | `=MINUTE(time)` | Extract minute |
| SECOND | `=SECOND(time)` | Extract second |

### Operators

**Arithmetic:** `+` `-` `*` `/` (parentheses for grouping)

**Comparison:** `<` `>` `=` `<=` `>=` `<>` (returns TRUE/FALSE)

**Text:** `&` (concatenation, e.g., `=A1&" "&B1`)

### Cell References

- **Single cell**: `A1`, `B2`, `AA100` (supports multi-letter columns)
- **Ranges**: `A1:B5`, `AA1:AZ50`
- **Relative**: `A1` - Changes when copied
- **Absolute**: `$A$1` - Stays fixed when copied
- **Mixed**: `$A1` or `A$1` - Partially fixed

Press `F4` while editing a reference to cycle through reference types:
`A1` → `$A$1` → `A$1` → `$A1` → `A1`

### String Literals

Use double quotes for text in formulas:
- `="Hello World"`
- `=IF(A1>0, "Positive", "Negative")`
- `=A1&" "&B1` (concatenates with space)

### AutoSum

Select a cell below or to the right of numbers and press `Alt+=` to automatically insert a SUM formula.

### View All Formulas

Press `Ctrl+`` ` (backtick) to toggle between showing formula results and showing the actual formulas in all cells.

---

## Formula Language Server

VisiGrid treats formulas like code, providing IDE-like assistance.

### Autocomplete

When you type `=` and start typing a function name, suggestions appear automatically. Use:
- `Arrow Up/Down` to navigate suggestions
- `Tab` or `Enter` to accept
- `Escape` to dismiss

### Signature Help

When you type `(` after a function name, a tooltip appears showing:
- The function signature with all parameters
- The current parameter highlighted
- Updates as you type commas between arguments

### Error Validation

VisiGrid checks your formulas in real-time:
- Syntax errors appear as red messages below the formula bar
- Validates parenthesis matching, operators, and function syntax
- Shows specific error messages before you even press Enter

### Context-Sensitive Help (F1)

Press `F1` while editing a formula to see detailed help for the function at cursor position:
- Function name and syntax
- Full description
- Category

Press `Escape` to dismiss the help popup.

---

## Named Ranges

Named ranges let you give meaningful names to cell ranges for easier formulas.

### Defining a Named Range

1. Select the cells you want to name
2. Press `Ctrl+Shift+N`
3. Enter a name (e.g., "SalesData")
4. Press Enter

Now you can use this name in formulas: `=SUM(SalesData)`

### Go to Definition (F12)

When your cursor is on a named range in the formula bar, press `F12` to jump to the cells that define that range.

### Find All References (Shift+F12)

Select any cell and press `Shift+F12` to see all cells that reference it. A popup shows:
- List of cells containing formulas that depend on the current cell
- Preview of each formula
- Click or press Enter to navigate to a reference

### Rename Symbol (Ctrl+Shift+R)

To rename a named range across all formulas:
1. Put your cursor on the name in any formula
2. Press `Ctrl+Shift+R`
3. Enter the new name
4. Press Enter to apply

All formulas using that name are updated automatically.

---

## Lua Scripting

VisiGrid includes a sandboxed Lua console for automating tasks and manipulating data programmatically. All operations are transactional and fully undoable.

### Opening the Console

Press `Ctrl+Shift+L` to toggle the Lua console panel. On first open, the input is pre-filled with `examples` - press Enter to see available scripts.

```
┌─────────────────────────────────────────┐
│ > sheet:get("A1")                       │
│ 42                                      │
│ > for r = 1, 10 do                      │
│ ...   sheet:set("A" .. r, r * 2)        │
│ ... end                                 │
│ nil                                     │
│ ops: 10 | cells: 10 | time: 1.2ms       │
│ >                                       │
│         Ctrl+Enter to run · Undo reverses the entire script │
└─────────────────────────────────────────┘
```

### Console Commands

Type these directly in the input:

| Command | Description |
|---------|-------------|
| `help` | Show all commands and shortcuts |
| `examples` | List available example scripts |
| `example N` | Load example N into input |
| `clear` | Clear output history |

### Built-in Examples

| # | Name | Description |
|---|------|-------------|
| 1 | Fill Series | Fill A1:A10 with powers of 2 |
| 2 | Trim Whitespace | Remove leading/trailing spaces from column A |
| 3 | Find Duplicates | Find and report duplicate values in column A |
| 4 | Normalize Dates | Convert date formats to YYYY-MM-DD |
| 5 | Compare Columns | Find mismatches between columns A and B |
| 6 | Generate Multiplication Table | Create 10x10 table starting at A1 |
| 7 | Sum Column | Calculate sum and average of numbers in column A |

Load any example: `example 1` or `example "Fill Series"`

### Sheet API

All coordinates are **1-indexed** to match spreadsheet conventions.

**Cell Access:**

| Method | Description |
|--------|-------------|
| `sheet:get("A1")` | Get value using A1 notation (shorthand) |
| `sheet:set("A1", val)` | Set value using A1 notation (shorthand) |
| `sheet:get_value(row, col)` | Get typed value by row/col |
| `sheet:set_value(row, col, val)` | Set cell value |
| `sheet:get_formula(row, col)` | Get formula string or nil |
| `sheet:set_formula(row, col, f)` | Set formula |

**Sheet Info:**

| Method | Description |
|--------|-------------|
| `sheet:rows()` | Number of rows with data |
| `sheet:cols()` | Number of columns with data |
| `sheet:selection()` | Current selection (see below) |

**Selection API:**

```lua
local sel = sheet:selection()
print(sel.range)      -- "A1:C5"
print(sel.start_row)  -- 1
print(sel.start_col)  -- 1
print(sel.end_row)    -- 5
print(sel.end_col)    -- 3
```

**Range API (Bulk Read/Write):**

```lua
local r = sheet:range("A1:C5")

-- Read all values as 2D table
local data = r:values()
print(data[1][1])  -- value at A1
print(data[2][3])  -- value at C2

-- Write 2D table to range
r:set_values({
    {1, 2, 3},
    {4, 5, 6},
    {7, 8, 9}
})

-- Range info
r:rows()     -- → 5
r:cols()     -- → 3
r:address()  -- → "A1:C5"
```

**Transaction Control:**

| Method | Description |
|--------|-------------|
| `sheet:begin()` | Start transaction (noop, for clarity) |
| `sheet:rollback()` | Discard pending changes, returns count |
| `sheet:commit()` | Commit changes (noop, auto-commits at end) |

### Example Scripts

```lua
-- Fill column with sequence
for i = 1, 100 do
    sheet:set("A" .. i, i * 2)
end
print("Filled 100 cells")
```

```lua
-- Process current selection
local sel = sheet:selection()
for row = sel.start_row, sel.end_row do
    for col = sel.start_col, sel.end_col do
        local val = sheet:get_value(row, col)
        if type(val) == "number" then
            sheet:set_value(row, col, val * 1.1)  -- +10%
        end
    end
end
```

```lua
-- Bulk transform with Range
local r = sheet:range("A1:C10")
local data = r:values()

for i, row in ipairs(data) do
    for j, val in ipairs(row) do
        if type(val) == "number" then
            data[i][j] = val * 2
        end
    end
end

r:set_values(data)
```

```lua
-- Conditional rollback
sheet:begin()
for i = 1, 100 do
    sheet:set("A" .. i, math.random(1, 100))
end

local sum = 0
for i = 1, 100 do
    sum = sum + (sheet:get("A" .. i) or 0)
end

if sum > 5000 then
    local discarded = sheet:rollback()
    print("Rolled back " .. discarded .. " ops (sum too high)")
else
    print("Sum: " .. sum)
end
```

### Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Toggle console | `Ctrl+Shift+L` |
| Execute script | `Enter` or `Ctrl+Enter` |
| Newline (multiline) | `Shift+Enter` |
| Previous command | `Up Arrow` |
| Next command | `Down Arrow` |
| Scroll output up | `Page Up` |
| Scroll output down | `Page Down` |
| Scroll to top | `Ctrl+Home` |
| Scroll to bottom | `Ctrl+End` |
| Clear output | `Ctrl+L` |
| Close console | `Escape` |

### Execution Stats

After each script runs, the console shows:
```
ops: 100 | cells: 50 | time: 12.5ms
```

- **ops**: Total operations queued
- **cells**: Unique cells modified (deduplicated)
- **time**: Wall-clock execution time

### Sandboxing and Safety

**Enabled Libraries:** `table`, `string`, `math`, `utf8`

**Disabled (for security):** `os`, `io`, `debug`, `package`, `require`, `load`, `loadfile`, `dofile`

**Resource Limits:**

| Limit | Value |
|-------|-------|
| Operations | 1,000,000 |
| Output lines | 5,000 |
| Instructions | 100,000,000 |
| Wall-clock | 30 seconds |

**Safety Guarantees:**
- No filesystem or network access
- Single undo entry per script (Ctrl+Z reverses entire script)
- Deterministic execution (same input = same output)

---

## Command Palette

Press `Ctrl+Shift+P` to open the command palette - a fuzzy-searchable list of all available commands.

### Search Modes (Prefixes)

| Prefix | Mode | Example |
|--------|------|---------|
| (none) | Commands + recent files | `save` |
| `>` | Commands only | `>bold` |
| `@` | Search cell contents | `@revenue` |
| `:` | Go to cell | `:A25` |
| `=` | Search functions | `=SUM` |
| `#` | Search settings | `#font` |

### Using the Palette

1. Type to filter results
2. Use `Arrow Up/Down` to navigate
3. Press `Enter` to execute/select
4. Press `Escape` to close

The command palette shows keyboard shortcuts next to commands, helping you learn them over time.

---

## Find and Search

### Find in Cells (Ctrl+F)

Press `Ctrl+F` to open the find dialog. Type your search term to find cells containing that text.

### Search via Command Palette

Use `@` prefix in the command palette to search cell contents:
- `Ctrl+Shift+P` then `@budget` finds all cells containing "budget"

---

## Cell Inspector

Press `Ctrl+Shift+I` to toggle the Cell Inspector panel. It shows detailed information about the selected cell:

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
└─────────────────────────────────┘
```

- **Precedents**: Cells that this formula depends on
- **Dependents**: Other cells that reference this cell

Click on any cell reference to navigate to it.

---

## Problems Panel

Press `Ctrl+Shift+M` to toggle the Problems panel. It shows all formula errors in your sheet:

```
Problems (3)
  D5: #REF! - Invalid cell reference
  E10: #DIV/0! - Division by zero
  F2: #NAME? - Unknown function
```

Click on any problem to navigate to that cell. The panel updates automatically as you fix errors.

---

## View Modes

### Zen Mode (F11)

Press `F11` to enter Zen Mode - a distraction-free view that hides:
- Menu bar
- Formula bar
- Format bar
- All panels

Press `Escape` or `F11` again to exit.

### Split View

| Action | Shortcut |
|--------|----------|
| Toggle split view | `Ctrl+\` |
| Switch between panes | `Ctrl+W` |

Split view shows two views of your sheet side by side. Each pane can scroll independently.

---

## Formatting

### Text Formatting

| Action | Shortcut |
|--------|----------|
| Bold | `Ctrl+B` |
| Italic | `Ctrl+I` |
| Underline | `Ctrl+U` |
| Strikethrough | `Ctrl+5` |

You can also use the format bar buttons or the Format menu. Formatting is applied to all selected cells.

### Format Painter

Copy formatting from one cell and apply it to others:

| Action | Shortcut |
|--------|----------|
| Copy format | `Ctrl+Shift+C` |
| Paste format | `Ctrl+Shift+V` |

**Usage:**
1. Select a cell with the formatting you want to copy
2. Press `Ctrl+Shift+C` (status bar confirms "● Format copied from A1")
3. Select one or more destination cells
4. Press `Ctrl+Shift+V` to apply the format

All formatting properties are copied: bold, italic, underline, strikethrough, alignment, number format, and text overflow mode.

### Horizontal Alignment

Use the Format menu or format bar to set text alignment:

| Alignment | Description |
|-----------|-------------|
| Left | Text aligns to the left edge (default) |
| Center | Text is centered in the cell |
| Right | Text aligns to the right edge |

### Vertical Alignment

| Alignment | Description |
|-----------|-------------|
| Top | Text aligns to the top of the cell |
| Middle | Text is centered vertically (default) |
| Bottom | Text aligns to the bottom of the cell |

Access via Format menu or Cell Inspector (`Ctrl+1`).

### Number Formats

| Format | Shortcut | Example |
|--------|----------|---------|
| General | `Ctrl+Shift+~` | 1234.56 (auto-detected) |
| Number | `Ctrl+Shift+!` | 1,234.56 |
| Currency | `Ctrl+Shift+$` | $1,234.56 |
| Percent | `Ctrl+Shift+%` | 12.34% |

To adjust decimal places:
- **Increase decimals**: Use format bar `.00→.000` button
- **Decrease decimals**: Use format bar `.000→.00` button

### Text Overflow

Control how text behaves when it's wider than the cell:

| Mode | Behavior |
|------|----------|
| Clip | Text is cut off at the cell boundary (default) |
| Wrap | Text wraps to multiple lines within the cell |
| Spill | Text overflows into adjacent empty cells |

**Spill behavior:**
- Spill stops at the next non-empty cell
- Spill stops at cells with visual formatting (bold, italic, etc.)
- Maximum spill is 8 cells

Access text overflow via the Cell Inspector (`Ctrl+1`).

### Column Width

| Action | Method |
|--------|--------|
| Resize column | Drag column border |
| Auto-fit column | Double-click column border |
| Auto-size column | `Ctrl+0` |

### Formatting Shortcuts Summary

| Action | Shortcut |
|--------|----------|
| Bold | `Ctrl+B` |
| Italic | `Ctrl+I` |
| Underline | `Ctrl+U` |
| Strikethrough | `Ctrl+5` |
| Copy format | `Ctrl+Shift+C` |
| Paste format | `Ctrl+Shift+V` |
| General format | `Ctrl+Shift+~` |
| Number format | `Ctrl+Shift+!` |
| Currency format | `Ctrl+Shift+$` |
| Percent format | `Ctrl+Shift+%` |
| Format dialog | `Ctrl+1` |

---

## File Operations

| Action | Shortcut |
|--------|----------|
| New file | `Ctrl+N` |
| Open file | `Ctrl+O` |
| Save | `Ctrl+S` |
| Quick Open (recent files) | `Ctrl+P` |

### Quick Open

Press `Ctrl+P` to see a list of recent files. Type to filter, then press Enter to open.

### File Format

VisiGrid uses SQLite-based `.sheet` files. These are:
- Fast to save and load
- Crash-safe (atomic writes)
- Self-contained (no external dependencies)

VisiGrid can also import/export CSV files.

---

## CLI Tools

VisiGrid provides `visigrid-cli` for headless spreadsheet operations in scripts and pipelines.

### Quick Examples

```bash
# Sum a column from CSV
cat sales.csv | visigrid-cli calc -f csv --headers '=SUM(revenue)'

# Convert CSV to JSON
visigrid-cli convert data.csv -t json --headers

# List all supported functions
visigrid-cli list-functions
```

### calc - Evaluate Formulas

Evaluate spreadsheet formulas against piped data:

```bash
visigrid-cli calc -f <format> [options] '<formula>'
```

| Option | Description |
|--------|-------------|
| `-f, --from` | Input format: `csv`, `tsv`, `json`, `lines`, `xlsx` |
| `--headers` | First row is headers (excluded from formulas) |
| `--into` | Load data starting at cell (default: A1) |
| `--delimiter` | CSV delimiter (default: comma) |
| `--spill` | Output format for array results: `csv` or `json` |

**Examples:**

```bash
# Average of column B
cat data.csv | visigrid-cli calc -f csv '=AVERAGE(B:B)'

# Sum with headers
echo -e "amount\n10\n20\n30" | visigrid-cli calc -f csv --headers '=SUM(amount)'

# Count lines in a file
cat file.txt | visigrid-cli calc -f lines '=COUNTA(A:A)'

# Array formula with spill
cat data.csv | visigrid-cli calc -f csv --spill json '=FILTER(A:A, B:B>10)'
```

### convert - Transform Formats

Convert between file formats:

```bash
visigrid-cli convert [input] -t <format> [options]
```

| Option | Description |
|--------|-------------|
| `-f, --from` | Input format (required when reading from stdin) |
| `-t, --to` | Output format: `csv`, `tsv`, `json`, `lines` |
| `-o, --output` | Output file (default: stdout) |
| `--headers` | First row is headers (affects JSON object keys) |
| `--delimiter` | CSV/TSV delimiter |

**Examples:**

```bash
# CSV to JSON with headers as keys
visigrid-cli convert data.csv -t json --headers

# JSON to CSV
curl api.example.com/data | visigrid-cli convert -f json -t csv

# Pipe CSV to JSON
cat data.csv | visigrid-cli convert -f csv -t json
```

### list-functions - Show Available Functions

```bash
visigrid-cli list-functions
```

Outputs all 96+ supported functions, one per line.

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Evaluation error (formula returned error) |
| 2 | Invalid arguments |
| 3 | I/O error |
| 4 | Parse error (malformed input) |
| 5 | Format error (unsupported format) |

See [docs/cli-v1.md](cli-v1.md) for complete specification.

### Session Restore

VisiGrid automatically saves your session on quit and restores it on next launch:

```bash
visigrid                    # Restore previous session
visigrid --no-restore       # Start fresh
visigrid -n                 # Start fresh (short form)
visigrid file.sheet         # Open specific file (skip session)
```

Session includes: current file, scroll position, selection, panel states, theme.

### Workspaces

VisiGrid supports per-project workspaces. Create a `.visigrid` marker file in your project root:

```bash
touch /path/to/project/.visigrid
```

When you open VisiGrid from within that directory, it loads/saves a project-specific session instead of the global one.

---

## Settings and Customization

VisiGrid uses text files for configuration, making settings version-controllable and shareable.

### Configuration Directory

Settings are stored in:
- Linux: `~/.config/visigrid/`
- macOS: `~/Library/Application Support/visigrid/`
- Windows: `%APPDATA%\visigrid\`

### settings.json

General application settings:

```json
{
  "editor.fontSize": 13,
  "editor.vimMode": false,
  "grid.defaultColumnWidth": 100,
  "grid.rowHeight": 24,
  "formula.autoRecalc": true,
  "file.recentFilesLimit": 10,
  "ui.showFormulaBar": true,
  "ui.showStatusBar": true
}
```

| Setting | Description |
|---------|-------------|
| `editor.fontSize` | Font size in pixels |
| `editor.vimMode` | Enable Vim-style navigation |
| `grid.defaultColumnWidth` | Default column width |
| `grid.rowHeight` | Row height in pixels |
| `formula.autoRecalc` | Auto-recalculate formulas |
| `file.recentFilesLimit` | Max recent files to remember |
| `ui.showFormulaBar` | Show/hide formula bar |
| `ui.showStatusBar` | Show/hide status bar |

### keybindings.json

Custom keyboard shortcuts:

```json
[
  { "key": "Ctrl+G", "command": "navigate.goto" },
  { "key": "Ctrl+Shift+N", "command": "namedRange.define" }
]
```

### Opening Settings

Use the command palette (`Ctrl+Shift+P`):
- "Open Settings" - Opens settings.json
- "Open Keyboard Shortcuts" - Opens keybindings.json

Or search settings directly with `#` prefix: `#column` shows column-related settings.

---

## Quick Reference

### Essential Shortcuts

| Action | Shortcut |
|--------|----------|
| Command Palette | `Ctrl+Shift+P` |
| Quick Open | `Ctrl+P` |
| Save | `Ctrl+S` |
| Undo/Redo | `Ctrl+Z` / `Ctrl+Y` |
| Copy/Cut/Paste | `Ctrl+C` / `Ctrl+X` / `Ctrl+V` |
| Find | `Ctrl+F` |
| Go to Cell | `Ctrl+G` |

### Formula Editing

| Action | Shortcut |
|--------|----------|
| Edit cell | `F2` |
| Context help | `F1` |
| Cycle reference type | `F4` |
| AutoSum | `Alt+=` |
| Toggle formula view | `Ctrl+`\` |

### Named Ranges

| Action | Shortcut |
|--------|----------|
| Define named range | `Ctrl+Shift+N` |
| Go to definition | `F12` |
| Find all references | `Shift+F12` |
| Rename symbol | `Ctrl+Shift+R` |

### Panels

| Panel | Shortcut |
|-------|----------|
| Cell Inspector | `Ctrl+Shift+I` |
| Problems | `Ctrl+Shift+M` |
| Lua REPL | `Ctrl+Shift+L` |
| Zen Mode | `F11` |
| Split View | `Ctrl+\` |

### Formatting

| Action | Shortcut |
|--------|----------|
| Bold | `Ctrl+B` |
| Italic | `Ctrl+I` |
| Underline | `Ctrl+U` |
| Strikethrough | `Ctrl+5` |
| Copy format | `Ctrl+Shift+C` |
| Paste format | `Ctrl+Shift+V` |
| Currency format | `Ctrl+Shift+$` |
| Percent format | `Ctrl+Shift+%` |
| Format dialog | `Ctrl+1` |

### Data Entry

| Action | Shortcut |
|--------|----------|
| Insert date | `Ctrl+;` |
| Insert time | `Ctrl+Shift+;` |
| Fill down | `Ctrl+D` |
| Fill right | `Ctrl+R` |
| Insert row/column | `Ctrl+Shift+=` |
| Delete row/column | `Ctrl+-` |
| Paste Special | `Ctrl+Alt+V` |

### Navigation

| Action | Shortcut |
|--------|----------|
| Keyboard hints | `g` (or `f` in Vim mode) |
| Open link in cell | `Ctrl+Click` |
| Switch split pane | `Ctrl+W` |

---

## Tips and Tricks

### 1. Learn Commands via Palette
The command palette (`Ctrl+Shift+P`) shows shortcuts next to each command. Use it to discover and learn shortcuts.

### 2. Use Named Ranges
Instead of `=SUM(B2:B50)`, define a named range "MonthlySales" and write `=SUM(MonthlySales)`. It's more readable and easier to maintain.

### 3. Check the Problems Panel
Keep the Problems panel open (`Ctrl+Shift+M`) while working with complex formulas. It catches errors early.

### 4. Use the Cell Inspector
When debugging formulas, the Cell Inspector (`Ctrl+Shift+I`) shows you exactly what cells a formula depends on and what other cells use its value.

### 5. Keyboard-First Workflow
VisiGrid is designed for keyboard-first usage. Learn a few shortcuts each day and your productivity will increase significantly.

### 6. Try Vim Mode
If you're familiar with Vim, enable `editor.vimMode` in settings for hjkl navigation. Press `i` to edit cells, `Escape` to navigate.

### 7. Automate with Lua
Use the Lua console (`Ctrl+Shift+L`) to automate repetitive tasks. Type `examples` to see 7 ready-to-run scripts. Use `sheet:selection()` to process the current selection, or `sheet:range("A1:C10")` for bulk operations. Every script is a single undo.

### 8. Use Paste Special
When pasting data, `Ctrl+Alt+V` lets you paste just values (no formulas), apply math operations, or transpose rows/columns.

---

## Getting Help

- Press `F1` while editing a formula for context-sensitive help
- Use the command palette (`Ctrl+Shift+P`) to find any command
- Search settings with `#` prefix in the command palette
- Check the Help menu for additional resources
