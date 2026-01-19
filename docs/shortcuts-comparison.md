# Keyboard Shortcuts: VisiGrid vs Excel (Windows)

This document compares VisiGrid's keyboard shortcuts against Microsoft Excel for Windows.

---

## Summary

| Category | VisiGrid | Excel | Coverage |
|----------|----------|-------|----------|
| Menu Accelerators (Alt) | 8 | 30+ | 27% |
| File Operations | 4 | 8 | 50% |
| Navigation | 12 | 25+ | 48% |
| Selection | 10 | 20+ | 50% |
| Editing | 14 | 25+ | 56% |
| Formatting | 8 | 15+ | 53% |
| Data/Formulas | 6 | 15+ | 40% |
| View/Display | 7 | 10+ | 70% |
| Function Keys (F1-F12) | 7 | 20+ | 35% |
| Context Menu | 0 | 3 | 0% |
| Edit Mode (in-cell) | 6 | 15+ | 40% |
| Mouse+Keyboard | 6 | 10+ | 60% |
| AutoComplete | 3 | 4+ | 75% |
| **Total** | **~87** | **200+** | **~43%** |

---

## Menu Accelerators (Alt Keys)

Excel uses Alt key combinations to access menus without a mouse. This is a Windows UI pattern.

### Classic Menu Bar (Alt + Letter)

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| File menu | Alt+F | Alt+F | ✅ |
| Edit menu | Alt+E | Alt+E | ✅ |
| View menu | Alt+V | Alt+V | ✅ |
| Insert menu | Alt+I | Alt+I | ✅ |
| Format menu | Alt+O | Alt+O | ✅ |
| Data menu | Alt+D | Alt+D | ✅ |
| Help menu | Alt+H | Alt+H | ✅ |
| Close menu / cancel | Escape | Escape | ✅ |

### Ribbon Access (Modern Excel)

Modern Excel uses Alt to show "KeyTips" - letters overlaid on ribbon buttons.

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Show KeyTips | Alt | — | ❌ |
| Home tab | Alt+H | — | ❌ |
| Insert tab | Alt+N | — | ❌ |
| Page Layout tab | Alt+P | — | ❌ |
| Formulas tab | Alt+M | — | ❌ |
| Data tab | Alt+A | — | ❌ |
| Review tab | Alt+R | — | ❌ |
| View tab | Alt+W | — | ❌ |

### Common Alt Sequences (Classic Menus)

| Action | Excel Sequence | VisiGrid | Status |
|--------|----------------|----------|--------|
| New | Alt+F, N | — | ❌ |
| Open | Alt+F, O | — | ❌ |
| Save | Alt+F, S | — | ❌ |
| Save As | Alt+F, A | — | ❌ |
| Print | Alt+F, P | — | ❌ |
| Undo | Alt+E, U | — | ❌ |
| Cut | Alt+E, T | — | ❌ |
| Copy | Alt+E, C | — | ❌ |
| Paste | Alt+E, P | — | ❌ |
| Paste Special | Alt+E, S | — | ❌ |
| Find | Alt+E, F | — | ❌ |
| Replace | Alt+E, E | — | ❌ |
| Go To | Alt+E, G | — | ❌ |
| Insert Row | Alt+I, R | — | ❌ |
| Insert Column | Alt+I, C | — | ❌ |
| Delete Row/Col | Alt+E, D | — | ❌ |
| Format Cells | Alt+O, E | — | ❌ |
| Column Width | Alt+O, C, W | — | ❌ |
| Row Height | Alt+O, R, E | — | ❌ |
| Sort | Alt+D, S | — | ❌ |
| Filter | Alt+D, F, F | — | ❌ |

### VisiGrid Approach

**Status: Menu accelerators implemented (Excel 2003 style)**
- Alt+F, Alt+E, Alt+V, Alt+I, Alt+O, Alt+D, Alt+H open their respective menus
- Underlined hotkey letters in menu bar (F̲ile, E̲dit, V̲iew, I̲nsert, Fo̲rmat, D̲ata, H̲elp)
- Sequential keypresses (Alt+F, then N for New) not yet implemented

The Command Palette (Ctrl+Shift+P) is the recommended modern alternative for power users.

---

## File Operations

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| New workbook | Ctrl+N | Ctrl+N | ✅ |
| Open file | Ctrl+O | Ctrl+O | ✅ |
| Save | Ctrl+S | Ctrl+S | ✅ |
| Save As | F12 | — | ❌ (F12 used for Go to Definition) |
| Print | Ctrl+P | — | ❌ (not applicable?) |
| Print Preview | Ctrl+F2 | — | ❌ |
| Close workbook | Ctrl+W | — | ❌ |
| Close Excel | Alt+F4 | — | ❌ (OS handles) |
| Quick Open | — | Ctrl+P | ✅ (VisiGrid unique) |

**Missing:** Print, Close workbook

---

## Navigation

### Cell Movement

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Move right | Tab | Tab | ✅ |
| Move left | Shift+Tab | Shift+Tab | ✅ |
| Move down | Enter | Enter | ✅ |
| Move up | Shift+Enter | Shift+Enter | ✅ |
| Move by arrow keys | Arrow keys | Arrow keys | ✅ |
| Jump to cell start | Ctrl+Home | Ctrl+Home | ✅ |
| Jump to last used cell | Ctrl+End | Ctrl+End | ✅ |
| Go to cell | Ctrl+G / F5 | Ctrl+G | ✅ (partial) |
| Jump to edge of data | Ctrl+Arrow | Ctrl+Arrow | ✅ |
| Page down | Page Down | Page Down | ✅ |
| Page up | Page Up | Page Up | ✅ |
| Move to row 1 | Ctrl+Up | — | ❌ |
| Move to column A | Ctrl+Left | — | ❌ |

### Scroll Without Moving Selection

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Scroll up one row | Ctrl+Up (scroll lock) | — | ❌ |
| Scroll down one row | Ctrl+Down (scroll lock) | — | ❌ |
| Scroll left one col | Ctrl+Left (scroll lock) | — | ❌ |
| Scroll right one col | Ctrl+Right (scroll lock) | — | ❌ |

### Sheet Navigation

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Next sheet | Ctrl+Page Down | Ctrl+Page Down | ✅ |
| Previous sheet | Ctrl+Page Up | Ctrl+Page Up | ✅ |
| Add new sheet | — | Click + button | ✅ |
| Move to next workbook | Ctrl+Tab | — | ❌ |

**Missing:** Scroll lock scrolling

---

## Selection

### Basic Selection

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Extend selection by arrow | Shift+Arrow | Shift+Arrow | ✅ |
| Extend to edge of data | Ctrl+Shift+Arrow | Ctrl+Shift+Arrow | ✅ |
| Select all | Ctrl+A | Ctrl+A | ✅ |
| Select entire column | Ctrl+Space | Ctrl+Space | ✅ |
| Select entire row | Shift+Space | Shift+Space | ✅ |
| Select to beginning | Ctrl+Shift+Home | — | ❌ |
| Select to end | Ctrl+Shift+End | Ctrl+Shift+End | ✅ |
| Select current region | Ctrl+Shift+* | — | ❌ |

### Special Selection

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Select cells with comments | Ctrl+Shift+O | — | ❌ |
| Select cells with formulas | — | — | ❌ |
| Select visible cells only | Alt+; | — | ❌ |
| Add to selection | Ctrl+Click | Ctrl+Click | ✅ |
| Select non-contiguous | Shift+F8 | — | ❌ |

### Named Ranges / Go To

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Name box (type cell ref) | F5 / Ctrl+G | Ctrl+G | ✅ |
| Define name | Ctrl+F3 | Ctrl+Shift+N | ✅ |
| Go to definition | — | F12 | ✅ |
| Rename named range | — | Ctrl+Shift+R | ✅ |
| Find all references | — | Shift+F12 | ✅ |
| Paste name | F3 | — | ❌ |

**Missing:** Select current region, visible cells only, paste name list

---

## Editing

### Clipboard

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Copy | Ctrl+C | Ctrl+C | ✅ |
| Cut | Ctrl+X | Ctrl+X | ✅ |
| Paste | Ctrl+V | Ctrl+V | ✅ |
| Paste Special | Ctrl+Alt+V | — | ❌ |
| Copy as picture | — | — | ❌ |

### Undo/Redo

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Undo | Ctrl+Z | Ctrl+Z | ✅ |
| Redo | Ctrl+Y | Ctrl+Y | ✅ |
| Repeat last action | F4 / Ctrl+Y | — | ❌ |

### Cell Editing

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Edit cell (in cell) | F2 | F2 | ✅ |
| Edit cell (formula bar) | — | — | ❌ |
| Delete cell contents | Delete | Delete | ✅ |
| Delete cell (shift) | Ctrl+- | Ctrl+- | ✅ |
| Insert cell | Ctrl+Shift+= | Ctrl+Shift+= | ✅ |
| Clear all (contents+format) | — | — | ❌ |
| Fill down | Ctrl+D | Ctrl+D | ✅ |
| Fill right | Ctrl+R | Ctrl+R | ✅ |
| Fill selection with entry | Ctrl+Enter | Ctrl+Enter | ✅ |
| Insert line break in cell | Alt+Enter | — | ❌ |
| Cancel edit | Escape | Escape | ✅ |
| Confirm edit | Enter | Enter | ✅ |

### Find/Replace

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Find | Ctrl+F | Ctrl+F | ✅ |
| Find next | F3 / Shift+F4 | — | ❌ |
| Find previous | Shift+F3 | — | ❌ |
| Replace | Ctrl+H | — | ❌ |

### Date/Time

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Insert current date | Ctrl+; | Ctrl+; | ✅ |
| Insert current time | Ctrl+Shift+; | Ctrl+Shift+; | ✅ |

**Missing:** Paste Special, Replace, line break in cell

---

## Formatting

### Text Formatting

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Bold | Ctrl+B | Ctrl+B | ✅ |
| Italic | Ctrl+I | Ctrl+I | ✅ |
| Underline | Ctrl+U | Ctrl+U | ✅ |
| Strikethrough | Ctrl+5 | Ctrl+5 | ✅ |

### Number Formatting

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| General format | Ctrl+Shift+~ | Ctrl+Shift+~ | ✅ |
| Currency format | Ctrl+Shift+$ | Ctrl+Shift+$ | ✅ |
| Percentage format | Ctrl+Shift+% | Ctrl+Shift+% | ✅ |
| Scientific format | Ctrl+Shift+^ | — | ❌ |
| Date format | Ctrl+Shift+# | — | ❌ |
| Time format | Ctrl+Shift+@ | — | ❌ |
| Number with 2 decimals | Ctrl+Shift+! | Ctrl+Shift+! | ✅ |

### Cell Formatting

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Format cells dialog | Ctrl+1 | — | ❌ |
| Add border | Ctrl+Shift+& | — | ❌ |
| Remove border | Ctrl+Shift+_ | — | ❌ |
| Apply outline border | — | — | ❌ |

### Alignment

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Align left | — | — | ❌ (menu only) |
| Align center | — | — | ❌ (menu only) |
| Align right | — | — | ❌ (menu only) |
| Indent | Ctrl+Alt+Tab | — | ❌ |

**Missing:** Scientific/Date/Time format shortcuts, Format Cells dialog, borders

---

## Data & Formulas

### Formula Entry

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| AutoSum | Alt+= | Alt+= | ✅ |
| Insert function | Shift+F3 | — | ❌ |
| Toggle formula view | Ctrl+` | Ctrl+` | ✅ |
| Cycle cell reference (F4) | F4 | F4 | ✅ |
| Function autocomplete | — | Auto (type =) | ✅ |
| Signature help | — | Auto (inside parens) | ✅ |
| Context help | — | F1 | ✅ |
| Find all references | — | Shift+F12 | ✅ |
| Array formula (legacy) | Ctrl+Shift+Enter | — | ❌ |
| Calculate workbook | F9 | F9 | ✅ |
| Calculate sheet | Shift+F9 | — | ❌ |
| Expand/collapse formula bar | Ctrl+Shift+U | — | ❌ |

### Data Operations

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Create table | Ctrl+T | — | ❌ |
| Toggle AutoFilter | Ctrl+Shift+L | — | ❌ |
| Sort ascending | — | — | ❌ |
| Sort descending | — | — | ❌ |
| Group rows/columns | Alt+Shift+Right | — | ❌ |
| Ungroup rows/columns | Alt+Shift+Left | — | ❌ |
| Flash Fill | Ctrl+E | — | ❌ |

### Comments/Notes

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Insert comment | Shift+F2 | — | ❌ |
| Edit comment | Shift+F2 | — | ❌ |
| Next comment | Ctrl+Shift+O | — | ❌ |

**Missing:** Insert function dialog, array formula, calculate, table creation, AutoFilter, Flash Fill, comments

---

## View & Display

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Toggle formula view | Ctrl+` | Ctrl+` | ✅ |
| Zoom in | Ctrl++ / Ctrl+= | Ctrl++ / Ctrl+= | ✅ |
| Zoom out | Ctrl+- | — | ⚠️ (Ctrl+- used for delete) |
| Zoom 100% | Ctrl+0 | Ctrl+0 | ✅ |
| Full screen | — | F11 | ✅ |
| Split panes | — | Ctrl+\ | ✅ |
| Switch split pane | — | Ctrl+W | ✅ |
| Freeze panes | — | — | ❌ |
| Hide column | Ctrl+0 | — | ❌ (Ctrl+0 used for zoom reset) |
| Unhide column | Ctrl+Shift+0 | — | ❌ |
| Hide row | Ctrl+9 | — | ❌ |
| Unhide row | Ctrl+Shift+9 | — | ❌ |
| Ribbon toggle | Ctrl+F1 | — | ❌ (no ribbon) |

**Missing:** Freeze panes, Hide/Unhide rows/columns
**Note:** Zoom out via Ctrl+- conflicts with delete row/col; use View menu or command palette

---

## VisiGrid Unique Shortcuts

These are VisiGrid-specific, inspired by code editors:

| Action | Shortcut | Inspiration |
|--------|----------|-------------|
| Command Palette | Ctrl+Shift+P | VS Code |
| Quick Open | Ctrl+P | VS Code |
| Problems Panel | Ctrl+Shift+M | VS Code |
| Cell Inspector | Ctrl+Shift+I | VS Code (DevTools) |
| Split View | Ctrl+\ | VS Code |
| Switch Split | Ctrl+W | — |
| Zen Mode | F11 | VS Code |
| Context Help | F1 | VS Code (hover docs alternative) |
| Find All References | Shift+F12 | VS Code |
| Rename Symbol | Ctrl+Shift+R | VS Code |
| Go to Definition | F12 | VS Code |
| Define Named Range | Ctrl+Shift+N | VS Code (new file) |

---

## Priority Implementation List

### Tier 1: Expected by Excel Users (High Impact)

| Shortcut | Action | Difficulty |
|----------|--------|------------|
| Ctrl+H | Replace | Medium |
| Ctrl+1 | Format cells dialog | Medium |
| Alt+Enter | Line break in cell | Medium |
| Ctrl+Shift+L | Toggle AutoFilter | Hard |

### Tier 2: Power User Shortcuts

| Shortcut | Action | Difficulty |
|----------|--------|------------|
| Ctrl+T | Create table | Hard |
| Shift+F3 | Insert function dialog | Medium |
| Ctrl+Shift+# | Date format | Easy |

### Tier 3: Nice to Have

| Shortcut | Action | Difficulty |
|----------|--------|------------|
| Shift+F2 | Insert comment | Medium |
| Ctrl+E | Flash Fill | Hard |
| Ctrl+Shift+* | Select current region | Medium |
| Alt+; | Select visible cells only | Medium |

---

## Recent Additions

The following shortcuts were recently implemented:

| Shortcut | Action | Notes |
|----------|--------|-------|
| Alt+F/E/V/I/O/D/H | Open menus | Excel 2003 style with underlined hotkeys |
| Ctrl+Shift+$/% | Currency/Percent format | Excel-compatible |
| Ctrl+Shift+~/! | General/Number format | Excel-compatible |
| Ctrl+5 | Strikethrough | Uses Unicode combining character |
| F9 | Recalculate | Force formula recalculation |
| Ctrl+= / Ctrl++ | Zoom in | Up to 300% |
| Ctrl+0 | Zoom reset | Returns to 100% |
| Ctrl+Enter | Fill selection | Applies value to all selected cells |
| Ctrl+PageDown | Next sheet | Multi-sheet navigation |
| Ctrl+PageUp | Previous sheet | Multi-sheet navigation |

---

## Implementation Notes

### Number Format Shortcuts

**Implemented:**
```
Ctrl+Shift+~  → General ✅
Ctrl+Shift+!  → Number (2 decimals) ✅
Ctrl+Shift+$  → Currency ($#,##0.00) ✅
Ctrl+Shift+%  → Percentage (0%) ✅
```

**Not yet implemented (need new NumberFormat variants):**
```
Ctrl+Shift+@  → Time (h:mm AM/PM)
Ctrl+Shift+#  → Date (d-mmm-yy)
Ctrl+Shift+^  → Scientific (0.00E+00)
```

### Multi-Sheet Support

Multi-sheet support is now implemented:
- Ctrl+Page Down (next sheet) ✅
- Ctrl+Page Up (previous sheet) ✅
- Click sheet tabs to switch ✅
- Click + button to add new sheet ✅
- Right-click sheet tab menu (not yet implemented)

### Conflict Notes

| Shortcut | Excel | VisiGrid | Resolution |
|----------|-------|----------|------------|
| Ctrl+P | Print | Quick Open | Keep VisiGrid (print less relevant) |
| Ctrl+W | Close workbook | Switch split | Keep VisiGrid |
| Ctrl+0 | Hide column | Zoom 100% | Keep VisiGrid (zoom more common) |
| Ctrl+- | Delete row/col | Zoom out | Keep delete (zoom out via menu) |

---

## Additional Categories (Not Yet Covered)

### Function Keys (F1-F12)

| Key | Excel Action | VisiGrid | Status |
|-----|--------------|----------|--------|
| F1 | Help | F1 (context help) | ✅ |
| F2 | Edit cell | F2 | ✅ |
| F3 | Paste name | — | ❌ |
| F4 | Cycle reference / Repeat | F4 | ✅ (cycle only) |
| F5 | Go To dialog | — | ❌ (use Ctrl+G) |
| F6 | Switch panes | — | ❌ |
| F7 | Spelling check | — | ❌ |
| F8 | Extend selection mode | — | ❌ |
| F9 | Calculate all | F9 | ✅ |
| F10 | Show KeyTips (like Alt) | — | ❌ |
| F11 | Create chart | F11 (zen mode) | ⚠️ Conflict |
| F12 | Go to definition | F12 | ✅ |
| Shift+F2 | Insert/edit comment | — | ❌ |
| Shift+F3 | Insert function | — | ❌ |
| Shift+F10 | Context menu | — | ❌ |
| Shift+F11 | Insert new sheet | — | ❌ |
| Shift+F12 | Find all references | Shift+F12 | ✅ |
| Ctrl+F1 | Toggle ribbon | — | ❌ |
| Ctrl+F3 | Name Manager | — | ❌ |
| Ctrl+F4 | Close workbook | — | ❌ |
| Ctrl+F5 | Restore window size | — | ❌ |
| Ctrl+F9 | Minimize workbook | — | ❌ |
| Ctrl+F10 | Maximize workbook | — | ❌ |

### Context Menu

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Show context menu | Right-click | Right-click | ❌ |
| Show context menu (keyboard) | Shift+F10 | — | ❌ |
| Application/Menu key | Menu key | — | ❌ |

**Note:** VisiGrid has no right-click context menu currently.

### Edit Mode vs Navigation Mode

Excel has different shortcuts depending on whether you're editing a cell:

**In Edit Mode (after F2 or typing):**

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Move cursor left | Left arrow | Left arrow | ✅ |
| Move cursor right | Right arrow | Right arrow | ✅ |
| Move word left | Ctrl+Left | — | ❌ |
| Move word right | Ctrl+Right | — | ❌ |
| Move to start of cell text | Home | — | ❌ |
| Move to end of cell text | End | — | ❌ |
| Select to start of cell text | Shift+Home | — | ❌ |
| Select to end of cell text | Shift+End | — | ❌ |
| Select all in cell | Ctrl+A | — | ❌ |
| Delete char left | Backspace | Backspace | ✅ |
| Delete char right | Delete | Delete | ✅ |
| Delete word left | Ctrl+Backspace | — | ❌ |
| Delete word right | Ctrl+Delete | — | ❌ |
| Select all in cell | Ctrl+A | — | ❌ |
| Move to start of cell | Home | — | ❌ |
| Move to end of cell | End | — | ❌ |
| Select to start | Shift+Home | — | ❌ |
| Select to end | Shift+End | — | ❌ |
| New line in cell | Alt+Enter | — | ❌ |
| Accept and stay | Ctrl+Enter | — | ❌ |
| Accept and move right | Tab | Tab | ✅ |
| Cancel edit | Escape | Escape | ✅ |

### Formula Bar Shortcuts

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Expand formula bar | Ctrl+Shift+U | — | ❌ |
| Move to formula bar | F2 (then click) | — | ❌ |
| Select all in formula bar | Ctrl+A | — | ❌ |

### Name Box Shortcuts

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Select Name Box | Ctrl+G / F5 | Ctrl+G | ✅ |
| Create name from selection | Ctrl+Shift+F3 | — | ❌ |
| Define name | Ctrl+F3 | — | ❌ |

### Mouse + Keyboard Combinations

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Add to selection | Ctrl+Click | Ctrl+Click | ✅ |
| Extend selection | Shift+Click | Shift+Click | ✅ |
| Select column | Click header | Click header | ✅ |
| Select row | Click row number | Click row number | ✅ |
| Auto-fill drag | Drag fill handle | — | ❌ |
| Copy while dragging | Ctrl+Drag | — | ❌ |
| Insert copied cells | Ctrl+Shift+Drag | — | ❌ |
| Resize column | Drag border | Drag border | ✅ |
| Auto-fit column | Double-click border | Double-click border | ✅ |
| Zoom | Ctrl+Scroll | — | ❌ |

### Workbook/Window Management

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| New window | Ctrl+N (new workbook) | Ctrl+N | ✅ |
| Switch windows | Ctrl+Tab | — | ❌ |
| Close window | Ctrl+W | — | ❌ |
| Minimize | — | — | ❌ (OS) |
| Maximize | — | — | ❌ (OS) |
| Move window | — | — | ❌ (OS) |

### Special Entry Shortcuts

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Enter same in all selected | Ctrl+Enter | Ctrl+Enter | ✅ |
| Enter as array formula | Ctrl+Shift+Enter | — | ❌ |
| Enter and move down | Enter | Enter | ✅ |
| Enter and move up | Shift+Enter | Shift+Enter | ✅ |
| Enter and move right | Tab | Tab | ✅ |
| Enter and stay | Ctrl+Enter | Ctrl+Enter | ✅ |

### AutoComplete / IntelliSense

| Action | Excel | VisiGrid | Status |
|--------|-------|----------|--------|
| Accept autocomplete | Tab | Tab/Enter | ✅ |
| Show function tooltip | Ctrl+Shift+A | F1 | ✅ (context help) |
| Show function arguments | Ctrl+A (in formula) | Auto (signature help) | ✅ |
| Cycle through suggestions | Arrow keys | Arrow keys | ✅ |

---

## Coverage Goals

**MVP Target:** 50% coverage of essential Excel shortcuts
- Focus on editing, navigation, basic formatting
- Defer: Print, multi-sheet, comments, Flash Fill

**v1.0 Target:** 70% coverage
- Add: Number format shortcuts, Replace, Format dialog
- Add: Sheet navigation (after multi-sheet)

**Long-term:** 85% coverage
- Add: Comments, Flash Fill, AutoFilter shortcuts
- Keep VisiGrid-unique editor-style shortcuts
