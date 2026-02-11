# VisiGrid Selection Semantics (v1)

> **This document defines v1 behavior. Changes require a version bump and are treated as breaking API changes. Do not modify without explicit approval.**
>
> **If observed behavior differs from this document, the behavior is a bug.**

The behavioral contract users rely on. Not implementation, not keybindings—the rules of the universe for how selection works.

---

## Definitions

### Active Cell
- Exactly one cell, always
- The anchor for keyboard movement
- Where edits occur
- Displayed with solid border
- Example: `B4`

### Primary Selection
- The selection range that contains the active cell
- Exactly one primary selection exists at all times
- Zero or more cells forming a rectangle
- When no range is explicitly selected, primary selection = active cell only
- Used for operations: copy, paste, fill, format, delete
- Example: `B4:E10` (with B4 as active cell)

### Additional Selections
- Zero or more independent rectangular ranges
- Created via Ctrl+Click or Ctrl+Drag
- Do not contain the active cell
- Cleared on any navigation or new selection
- Used for: applying formats, simultaneous edits (future)

### Edit State
- Whether the active cell is being edited
- Two sub-modes: **Edit** (text) and **Formula** (starts with `=`)
- Overrides navigation semantics
- Changes the meaning of arrow keys, Enter, Escape

---

## Navigation Semantics

### Arrow Keys (no modifiers)
1. Move active cell by one in that direction
2. Collapse selection to single cell
3. Clear additional selections
4. If in edit mode: move cursor within text (left/right) or commit and move (up/down)

### Ctrl+Arrow
1. Jump to next boundary (empty↔filled transition)
2. Collapse selection to single cell
3. Clear additional selections

### Tab / Shift+Tab
1. Move active cell right/left by one column
2. If editing: commit edit, then move
3. Collapse selection to single cell

### Enter / Shift+Enter
1. If editing: commit edit
2. Move active cell down/up by one row
3. Collapse selection to single cell

### Home / End
1. Home: move to column A in current row
2. End: move to last used column in current row
3. Collapse selection

### Ctrl+Home / Ctrl+End
1. Ctrl+Home: move to A1
2. Ctrl+End: move to last used cell in sheet
3. Collapse selection

---

## Selection Expansion Rules

### Shift+Arrow
1. Active cell (anchor) stays fixed
2. Extend selection edge in arrow direction
3. Selection is always rectangular
4. Additional selections are preserved

### Shift+Click
1. Active cell (anchor) stays fixed
2. Extend selection to clicked cell
3. Creates rectangle from anchor to click target
4. Clear additional selections

### Click (no modifiers)
1. Set active cell to clicked cell
2. Collapse selection to single cell
3. Clear additional selections

### Drag
1. First cell of drag becomes active cell
2. Selection extends to current mouse position
3. Clear additional selections
4. On release: selection is finalized

### Ctrl+Click
1. Active cell moves to clicked cell
2. Previous selection becomes an additional selection
3. Clicked cell becomes new primary selection (single cell)

> **Note:** Ctrl+Click always moves the active cell. This is an intentional simplification versus Excel, where Ctrl+Click adds/removes selections without necessarily moving the active cell. Our rule is cleaner for a keyboard-first tool.

### Ctrl+Drag
1. Previous selection becomes an additional selection
2. Drag creates new selection with new active cell

### Shift+Ctrl+Arrow
1. Extend selection to next boundary
2. Active cell stays fixed

---

## Edit Mode Transitions

### Entering Edit Mode

| Action | Result |
|--------|--------|
| Type any character | Enter Edit mode, replace cell content |
| F2 | Enter Edit mode, cursor at end of existing content |
| Double-click | Enter Edit mode, cursor at click position |
| Type `=` | Enter Formula mode |

### While in Edit Mode

| Action | Result |
|--------|--------|
| Arrow Left/Right | Move cursor within text |
| Arrow Up/Down | Commit edit, move active cell |
| Enter | Commit edit, move down |
| Shift+Enter | Commit edit, move up |
| Tab | Commit edit, move right |
| Shift+Tab | Commit edit, move left |
| Escape | Cancel edit, restore original value |
| Click another cell | Commit edit, select clicked cell |

### Formula Mode Special Behavior

| Action | Result |
|--------|--------|
| Arrow keys | Insert/move cell reference at cursor |
| Click cell | Insert cell reference at cursor |
| Shift+Click | Insert range reference (anchor to click) |
| Drag cells | Insert range reference |
| Enter | Commit formula, move down |
| Escape | Cancel formula entry |

### Formula Mode Invariants
- The active cell does not change during formula entry
- Selection changes are temporary and used only to construct references
- Navigation state resumes only after commit or cancel
- The formula bar always reflects the current formula text, not the referenced cells

---

## Mouse Interaction Rules

| Action | Context | Result |
|--------|---------|--------|
| Click | Normal | Select single cell |
| Click | While editing | Commit edit, select clicked cell |
| Click | While in formula | Insert reference (not commit) |
| Double-click | Normal | Enter edit mode |
| Double-click | On header | Auto-fit column/row |
| Drag | From cell | Select range |
| Drag | From header | Select entire rows/columns |
| Right-click | Any | Context menu (future) |

---

## Invariants (Always True)

1. There is exactly one active cell at all times
2. The active cell is always within the primary selection
3. Primary selection is always exactly one rectangle
4. Additional selections are zero or more non-overlapping rectangles
5. Edit mode is mutually exclusive with navigation
6. Arrow keys without Shift always collapse to single cell
7. Escape always cancels current operation and returns to navigation

---

## Explicit Non-Goals

We do **not** support:

- Non-rectangular selections (Excel's Ctrl+Click creates rectangles, not arbitrary shapes)
- Selection across multiple sheets
- Hidden rows/columns affecting selection geometry
- Merged cells (v1)
- Selection of objects (charts, images) - cells only
- "Extend mode" (F8 in Excel) - we use Shift consistently

---

## State Transitions Summary

```
┌─────────────┐
│ Navigation  │◄──────────────────────────────┐
└──────┬──────┘                               │
       │ type char / F2 / double-click        │ Escape / commit
       ▼                                      │
┌─────────────┐                               │
│   Editing   │───────────────────────────────┤
└──────┬──────┘                               │
       │ type '=' at start                    │
       ▼                                      │
┌─────────────┐                               │
│   Formula   │───────────────────────────────┘
└─────────────┘
       │
       │ arrow/click inserts references
       ▼
   (stays in Formula until commit/cancel)
```

---

## Version History

- **v1** (2026-01): Initial specification based on Excel compatibility
