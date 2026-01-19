# Keyboard Shortcuts: VisiGrid gpui Implementation

Current status of keyboard shortcuts in the gpui version.

---

## Summary

| Category | Implemented | Target | Coverage |
|----------|-------------|--------|----------|
| Navigation | 12 | 20 | 60% |
| Selection | 5 | 12 | 42% |
| Editing | 10 | 20 | 50% |
| Clipboard | 3 | 5 | 60% |
| File | 4 | 6 | 67% |
| Formatting | 3 | 10 | 30% |
| View | 0 | 8 | 0% |
| **Total** | **37** | **81** | **46%** |

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
| Ctrl+A | Select all | ✅ |
| Shift+Click | Extend to cell | ✅ |
| Double-click | Edit cell | ✅ |
| Click | Select cell | ✅ |

### Editing

| Shortcut | Action | Status |
|----------|--------|--------|
| F2 | Start edit | ✅ |
| Escape | Cancel edit | ✅ |
| Enter | Confirm edit | ✅ |
| Delete | Delete selection | ✅ |
| Backspace | Backspace in edit | ✅ |
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

---

## NOT YET Implemented

### Priority 0 (MVP Blockers)

| Shortcut | Action | Notes |
|----------|--------|-------|
| Ctrl+Shift+P | Command Palette | Core differentiator |
| Ctrl+D | Fill Down | Common workflow |
| Ctrl+R | Fill Right | Common workflow |
| Ctrl+Enter | Fill selection | Multi-edit |

### Priority 1 (Expected)

| Shortcut | Action | Notes |
|----------|--------|-------|
| Ctrl+Click | Add to selection | Discontiguous |
| Alt+F/E/V/O/D/H | Menu accelerators | Excel standard |
| Ctrl+H | Replace | Find & Replace |
| Ctrl+1 | Format Cells | Dialog |
| Ctrl+PageUp/Down | Sheet navigation | Multi-sheet |
| Ctrl+Shift+Arrow | Extend to edge | Selection |

### Priority 2 (Power User)

| Shortcut | Action | Notes |
|----------|--------|-------|
| Ctrl+Shift+$ | Currency format | |
| Ctrl+Shift+% | Percent format | |
| Ctrl+Shift+~ | General format | |
| Ctrl+Shift+! | Number format | |
| Ctrl+5 | Strikethrough | |
| Ctrl++ | Zoom in | |
| Ctrl+- | Zoom out | Conflicts with delete |
| Ctrl+0 | Zoom reset | |
| F9 | Recalculate | |
| F11 | Zen mode | |
| Alt+= | AutoSum | |

### Priority 3 (Nice to Have)

| Shortcut | Action | Notes |
|----------|--------|-------|
| Ctrl+P | Quick Open | VS Code style |
| Ctrl+Shift+M | Problems Panel | |
| Ctrl+Shift+I | Cell Inspector | |
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
| Ctrl+Shift+P | Command Palette | ❌ |
| Ctrl+P | Quick Open | ❌ |
| Ctrl+Shift+M | Problems Panel | ❌ |
| Ctrl+Shift+I | Cell Inspector | ❌ |
| Ctrl+\ | Split View | ❌ |
| F11 | Zen Mode | ❌ |
| F1 | Context Help | ❌ |
| F12 | Go to Definition | ❌ |
| Shift+F12 | Find All References | ❌ |
| Ctrl+Shift+N | Define Named Range | ❌ |
| Ctrl+Shift+R | Rename | ❌ |

---

## Implementation Order

### Sprint 1: Core Workflow
1. Ctrl+Shift+P (Command Palette)
2. Ctrl+D/R (Fill Down/Right)
3. Ctrl+Enter (Fill selection)
4. Ctrl+Click (Discontiguous selection)

### Sprint 2: Excel Compatibility
1. Alt menu accelerators
2. Ctrl+H (Replace)
3. Ctrl+1 (Format Cells)
4. Number format shortcuts

### Sprint 3: Multi-Sheet
1. Ctrl+PageUp/Down
2. Sheet tab clicks
3. Sheet context menu

### Sprint 4: Power Features
1. Zoom shortcuts
2. F9 recalculate
3. F11 zen mode
4. Alt+= AutoSum
