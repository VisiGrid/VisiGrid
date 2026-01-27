# VisiGrid Keyboard Shortcuts

Complete reference for keyboard shortcuts.

---

## Navigation

| Shortcut | Action |
|----------|--------|
| Arrow keys | Move selection |
| Tab | Move right |
| Shift+Tab | Move left |
| Enter | Confirm edit, move down |
| Shift+Enter | Confirm edit, move up |
| Ctrl+Home | Go to A1 |
| Ctrl+End | Go to last cell |
| Ctrl+Arrow | Jump to edge of data |
| Ctrl+Backspace | Jump to active cell (scroll view) |
| Page Up/Down | Page scroll |
| Ctrl+G | Go To dialog |
| Ctrl+F | Find |
| F3 | Find next |
| Shift+F3 | Find previous |

---

## Selection

| Shortcut | Action |
|----------|--------|
| Shift+Arrow | Extend selection |
| Ctrl+Shift+Arrow | Extend to edge of data |
| Ctrl+A | Select all |
| Shift+Click | Extend selection to cell |
| Ctrl+Click | Add to selection (discontiguous) |
| Shift+Space | Select entire row |
| Ctrl+Space | Select entire column |
| Ctrl+Shift+End | Select to last cell |

---

## Editing

| Shortcut | Action |
|----------|--------|
| F2 | Edit cell |
| Escape | Cancel edit |
| Enter | Confirm edit |
| Ctrl+Enter | Fill selection / Open link |
| Delete | Clear cell contents |
| Ctrl+D | Fill Down |
| Ctrl+R | Fill Right |
| Ctrl+Z | Undo |
| Ctrl+Y | Redo |
| Ctrl+Shift+Z | Redo (alternate) |
| Ctrl+; | Insert current date |
| Ctrl+Shift+; | Insert current time |
| F4 | Cycle cell reference ($A$1 → A$1 → $A1 → A1) |

---

## Clipboard

| Shortcut | Action |
|----------|--------|
| Ctrl+C | Copy |
| Ctrl+X | Cut |
| Ctrl+V | Paste |

---

## Formatting

| Shortcut | Action |
|----------|--------|
| Ctrl+B | Bold |
| Ctrl+I | Italic |
| Ctrl+U | Underline |
| Ctrl+5 | Strikethrough |
| Ctrl+1 | Format Cells dialog |
| Ctrl+Shift+~ | General format |
| Ctrl+Shift+! | Number format (2 decimals) |
| Ctrl+Shift+$ | Currency format |
| Ctrl+Shift+% | Percent format |

---

## Formulas

| Shortcut | Action |
|----------|--------|
| = | Start formula |
| Alt+= | AutoSum |
| Ctrl+` | Toggle formula view |
| F9 | Recalculate all |
| F1 | Context help (in formula) |
| Tab/Enter | Accept autocomplete |

---

## Named Ranges

| Shortcut | Action |
|----------|--------|
| Ctrl+Shift+N | Create named range |
| F12 | Go to definition |
| Shift+F12 | Find all references |
| Ctrl+Shift+R | Rename symbol |

---

## View

| Shortcut | Action |
|----------|--------|
| Ctrl+Shift+P | Command Palette |
| Ctrl+P | Fuzzy Finder (cells, ranges, files) |
| Ctrl+Shift+I | Inspector Panel |
| Ctrl+Shift+M | Problems Panel |
| Ctrl+Shift+L | Lua Console |
| Ctrl+K Ctrl+T | Theme Picker |
| Ctrl+, | Preferences |
| F11 | Zen Mode (hide UI) |
| Ctrl+\ | Split View |
| Ctrl+W | Switch split pane |
| Ctrl++ | Zoom in |
| Ctrl+0 | Zoom reset (100%) |

---

## File

| Shortcut | Action |
|----------|--------|
| Ctrl+N | New file |
| Ctrl+O | Open file |
| Ctrl+S | Save |
| Ctrl+Shift+S | Save As |

---

## Sheets

| Shortcut | Action |
|----------|--------|
| Ctrl+Page Down | Next sheet |
| Ctrl+Page Up | Previous sheet |
| Shift+F11 | Add new sheet |

---

## Alt Key (Excel-style Command Access)

VisiGrid supports Excel-style Alt navigation using a scoped command palette.

- Alt enters command mode
- First key selects a scope
- Commands are searched and executed via the palette

This preserves Excel muscle memory while keeping commands discoverable and explainable.

| Shortcut | Scope | What's in it |
|----------|-------|--------------|
| Alt+A | Data | Filter, Sort, Validation |
| Alt+E | Edit | Fill, Clear, Undo/Redo |
| Alt+F | File | Open, Save, Export |
| Alt+V | View | Inspector, Zoom, Freeze |
| Alt+T | Tools | Trace, Verified Mode, Explain |
| Alt+O | Format | Bold, Italic, Colors |
| Alt+H | Help | About, Shortcuts |
| Alt+Down | — | Open validation/filter dropdown |

**Example flow:** Alt+A → type "sort" → Enter

---

## Optional: Keyboard Hints (Vimium-style)

Enable in Preferences → Navigation.

| Shortcut | Action |
|----------|--------|
| g | Show cell labels |
| a-z, aa-zz | Jump to labeled cell |
| Escape | Cancel hint mode |
| Backspace | Delete last character |

---

## Optional: Vim Mode

Enable in Preferences → Navigation.

| Shortcut | Action |
|----------|--------|
| h/j/k/l | Move left/down/up/right |
| i | Enter edit mode |
| a | Enter edit mode (append) |
| 0 | Jump to column A |
| $ | Jump to last data column |
| w | Next filled cell (right) |
| b | Previous filled cell (left) |
| gg | Go to A1 |
| G | Go to bottom of data |
| Shift+hjkl | Extend selection |

---

## Keybinding Customization

Remap any shortcut via JSON config.

**Location:** `~/.config/visigrid/keybindings.json`

**Access:** Command Palette → "Open Keyboard Shortcuts"

**Example:**
```json
{
  "ctrl-shift-d": "edit.filldown",
  "ctrl-;": "edit.trim"
}
```

User keybindings override defaults.

---

## Keyboard Invariants

These shortcuts **never change** — they're Excel muscle memory that VisiGrid respects unconditionally. When using macOS with Ctrl preference, these remain Ctrl-based (not Cmd).

| Shortcut | Action | Rationale |
|----------|--------|-----------|
| F2 | Edit cell | Universal spreadsheet convention |
| F4 | Cycle reference (in formula) | $A$1 → A$1 → $A1 → A1 |
| F9 | Recalculate all | Excel's "refresh" key |
| Ctrl+D | Fill Down | Fundamental data entry |
| Ctrl+R | Fill Right | Fundamental data entry |
| Ctrl+Arrow | Jump to data edge | Power navigation |
| Ctrl+Shift+Arrow | Extend to data edge | Power selection |
| Ctrl+Backspace | Jump to active cell | View navigation |
| Alt+= | AutoSum | Quick sum formula |
| Alt+Down | Open dropdown | Validation list → Filter menu → nothing |

---

## Not Yet Implemented

| Shortcut | Action |
|----------|--------|
| Ctrl+Shift+V | Paste Special |
| Ctrl+H | Find and Replace |
| Ctrl+- | Zoom out |
| Alt+Enter | Line break in cell |
