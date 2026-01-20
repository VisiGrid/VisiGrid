# Keyboard Shortcuts: VisiGrid gpui Implementation

Current status of keyboard shortcuts in the gpui version.

---

## Summary

| Category | Implemented | Target | Coverage |
|----------|-------------|--------|----------|
| Navigation | 14 | 20 | 70% |
| Selection | 11 | 12 | 92% |
| Editing | 14 | 20 | 70% |
| Clipboard | 3 | 5 | 60% |
| File | 5 | 6 | 83% |
| Formatting | 3 | 10 | 30% |
| View | 5 | 8 | 63% |
| Menu | 7 | 7 | 100% |
| Sheets | 5 | 5 | 100% |
| Named Ranges | 3 | 3 | 100% |
| **Total** | **70** | **96** | **73%** |

---

## Implemented Shortcuts

### Navigation

| Shortcut | Action | Status |
|----------|--------|--------|
| Arrow keys | Move selection | ✅ |
| Tab | Move right | ✅ |
| Shift+Tab | Move left | ✅ |
| Enter | Confirm edit, move down | ✅ |
| Ctrl+Home | Go to A1 | ✅ |
| Ctrl+End | Go to last cell | ✅ |
| Ctrl+Arrow | Jump to edge of data | ✅ |
| Page Up/Down | Page scroll | ✅ |
| Ctrl+G | Go To dialog | ✅ |
| Ctrl+F | Find dialog | ✅ |
| F3 | Find next | ✅ |
| Shift+F3 | Find previous | ✅ |

### Selection

| Shortcut | Action | Status |
|----------|--------|--------|
| Shift+Arrow | Extend selection | ✅ |
| Ctrl+Shift+Arrow | Extend to edge of data | ✅ |
| Ctrl+A | Select all | ✅ |
| Shift+Click | Extend to cell | ✅ |
| Ctrl+Click | Add to selection | ✅ |
| Double-click | Edit cell | ✅ |
| Click | Select cell | ✅ |
| Shift+Space | Select row | ✅ |
| Ctrl+Space | Select column | ✅ |

### Editing

| Shortcut | Action | Status |
|----------|--------|--------|
| F2 | Start edit | ✅ |
| Escape | Cancel edit | ✅ |
| Enter | Confirm edit | ✅ |
| Ctrl+Enter | Confirm without moving / Multi-edit | ✅ |
| Delete | Delete selection | ✅ |
| Backspace | Backspace in edit | ✅ |
| Ctrl+D | Fill Down | ✅ |
| Ctrl+R | Fill Right | ✅ |
| Ctrl+Z | Undo | ✅ |
| Ctrl+Y | Redo | ✅ |
| Ctrl+Shift+Z | Redo (alt) | ✅ |
| Any character | Start edit with char | ✅ |

### Clipboard

| Shortcut | Action | Status |
|----------|--------|--------|
| Ctrl+C | Copy | ✅ |
| Ctrl+X | Cut | ✅ |
| Ctrl+V | Paste | ✅ |

### File

| Shortcut | Action | Status |
|----------|--------|--------|
| Ctrl+N | New file | ✅ |
| Ctrl+O | Open file | ✅ |
| Ctrl+S | Save | ✅ |
| Ctrl+Shift+S | Save As | ✅ |

### Formatting

| Shortcut | Action | Status |
|----------|--------|--------|
| Ctrl+B | Bold | ✅ |
| Ctrl+I | Italic | ✅ |
| Ctrl+U | Underline | ✅ |

### View

| Shortcut | Action | Status |
|----------|--------|--------|
| Ctrl+Shift+P | Command Palette | ✅ |
| Ctrl+P | Fuzzy Finder | ✅ |
| Ctrl+Shift+I | Inspector Panel | ✅ |
| Ctrl+K Ctrl+T | Theme Picker | ✅ |
| Ctrl+, | Preferences | ✅ |

### Named Ranges

| Shortcut | Action | Status |
|----------|--------|--------|
| Ctrl+Shift+N | Create Named Range | ✅ |
| Ctrl+Shift+R | Rename Symbol | ✅ |
| F2 (on name) | Edit Named Range | ✅ |

### Menu (Excel 2003 Style)

| Shortcut | Action | Status |
|----------|--------|--------|
| Alt+F | File menu | ✅ |
| Alt+E | Edit menu | ✅ |
| Alt+V | View menu | ✅ |
| Alt+I | Insert menu | ✅ |
| Alt+O | Format menu | ✅ |
| Alt+D | Data menu | ✅ |
| Alt+H | Help menu | ✅ |

### Sheet Navigation

| Shortcut | Action | Status |
|----------|--------|--------|
| Ctrl+Page Down | Next sheet | ✅ |
| Ctrl+Page Up | Previous sheet | ✅ |
| Shift+F11 | Add new sheet | ✅ |
| Click tab | Switch to sheet | ✅ |
| Click + | Add new sheet | ✅ |

---

## NOT YET Implemented

### Priority 1 (Expected)

| Shortcut | Action | Notes |
|----------|--------|-------|
| Ctrl+H | Replace | Find & Replace |
| Ctrl+1 | Format Cells | Dialog |

### Priority 2 (Power User)

| Shortcut | Action | Notes |
|----------|--------|-------|
| Ctrl+Shift+$ | Currency format | |
| Ctrl+Shift+% | Percent format | |
| Ctrl+Shift+~ | General format | |
| Ctrl+Shift+! | Number format | |
| Ctrl+5 | Strikethrough | |
| Ctrl++ | Zoom in | |
| Ctrl+- | Zoom out | |
| Ctrl+0 | Zoom reset | |
| F9 | Recalculate | |
| F11 | Zen mode | |
| Alt+= | AutoSum | |

### Priority 3 (Nice to Have)

| Shortcut | Action | Notes |
|----------|--------|-------|
| Ctrl+Shift+M | Problems Panel | |
| Ctrl+\ | Split view | |
| F1 | Context help | |
| F12 | Go to Definition | Named ranges |
| Shift+F12 | Find all references | |
| Ctrl+; | Insert date | |
| Ctrl+Shift+; | Insert time | |
| Alt+Enter | Line break in cell | |

---

## VisiGrid-Unique Shortcuts

Editor-inspired shortcuts that differentiate from Excel:

| Shortcut | Action | Status |
|----------|--------|--------|
| Ctrl+Shift+P | Command Palette | ✅ |
| Ctrl+P | Fuzzy Finder | ✅ |
| Ctrl+Shift+I | Cell Inspector | ✅ |
| Ctrl+K Ctrl+T | Theme Picker | ✅ |
| Ctrl+, | Preferences | ✅ |
| Ctrl+Shift+N | Define Named Range | ✅ |
| Ctrl+Shift+R | Rename Symbol | ✅ |
| Ctrl+Shift+M | Problems Panel | ❌ |
| Ctrl+\ | Split View | ❌ |
| F11 | Zen Mode | ❌ |
| F1 | Context Help | ❌ |
| F12 | Go to Definition | ❌ |
| Shift+F12 | Find All References | ❌ |

---

## Implementation Order

### Sprint 1: Core Workflow ✅ COMPLETE
1. ✅ Ctrl+Shift+P (Command Palette)
2. ✅ Ctrl+D/R (Fill Down/Right)
3. ✅ Ctrl+Enter (Multi-edit)
4. ✅ Ctrl+Click (Discontiguous selection)

### Sprint 2: Excel Compatibility ✅ COMPLETE
1. ✅ Alt menu accelerators (Alt+F/E/V/I/O/D/H)
2. ✅ Excel file import (xlsx/xls/xlsb/ods)
3. ❌ Ctrl+H (Replace)
4. ❌ Ctrl+1 (Format Cells)
5. ❌ Number format shortcuts

### Sprint 3: Multi-Sheet ✅ COMPLETE
1. ✅ Ctrl+PageUp/Down
2. ✅ Sheet tab clicks
3. ✅ Shift+F11 to add sheet
4. ✅ Sheet context menu (rename, delete)

### Sprint 4: Power Features ✅ COMPLETE
1. ✅ Ctrl+Shift+N (Named ranges)
2. ✅ Ctrl+Shift+R (Rename symbol)
3. ✅ Ctrl+P (Fuzzy finder)
4. ✅ Ctrl+Shift+I (Inspector panel)
5. ✅ Ctrl+K Ctrl+T (Theme picker)
6. ✅ Ctrl+, (Preferences)

### Sprint 5: Polish (Next)
1. Zoom shortcuts (Ctrl++/-)
2. F9 recalculate
3. F11 zen mode
4. Alt+= AutoSum
5. Ctrl+H (Replace)
