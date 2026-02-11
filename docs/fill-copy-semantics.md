# VisiGrid Fill & Copy Semantics (v1)

> **This document defines v1 behavior. Changes require a version bump and are treated as breaking API changes. Do not modify without explicit approval.**
>
> **If observed behavior differs from this document, the behavior is a bug.**

Fill & copy semantics define what happens when data is replicated, propagated, or inserted based on a selection. If selection semantics define *what is selected*, fill & copy semantics define *what happens to data when you act on that selection*.

---

## Definitions

### Source Range
- The cells providing data for an operation
- For Copy: the primary selection at time of copy
- For Fill Down: the first row of the primary selection
- For Fill Right: the first column of the primary selection

### Target Range
- The cells receiving data from an operation
- For Paste: rectangle starting at active cell, sized to match source
- For Fill Down: all rows below the first in primary selection
- For Fill Right: all columns after the first in primary selection

### Clipboard
- System clipboard (shared with OS)
- Internal representation: tab-separated values (TSV)
- Rows separated by newlines, columns by tabs
- Formulas stored as raw text (e.g., `=A1+B1`)

### Formula Reference Adjustment
- When a formula is copied/filled to a new location:
  - **Relative references** shift by the delta (e.g., `A1` → `A2` when filled down one row)
  - **Absolute references** do not shift (e.g., `$A$1` stays `$A$1`)
  - **Mixed references** shift only the relative component (e.g., `$A1` → `$A2`, `A$1` → `B$1`)

---

## Copy Operation

### Ctrl+C / Edit → Copy

**Source:** Primary selection only

**Behavior:**
1. Read all cells in primary selection
2. Convert to TSV format (tabs between columns, newlines between rows)
3. Write to system clipboard
4. Store internally for paste fallback

**Rules:**
- Additional selections are ignored
- Empty cells become empty strings in TSV
- Formulas are copied as raw formula text, not computed values
- Selection is preserved after copy (no collapse)
- The copy origin (top-left cell of the primary selection) is recorded at copy time and used for all subsequent paste delta calculations until the next copy or cut

### While Editing (Edit Mode Copy)

**Source:** Selected text in edit buffer (or all text if no selection)

**Behavior:**
1. Copy selected text (or entire edit value if no text selection)
2. Write to system clipboard

---

## Cut Operation

### Ctrl+X / Edit → Cut

**Behavior:**
1. Perform Copy operation
2. Clear all cells in primary selection
3. Record as undoable batch operation

**Rules:**
- Additional selections are ignored
- Selection is preserved after cut
- Undo restores both content and selection

---

## Paste Operation

### Ctrl+V / Edit → Paste

**Target:** Rectangle starting at active cell

**Behavior:**
1. Read from system clipboard (fall back to internal clipboard)
2. Parse as TSV (detect rows and columns)
3. Compute target rectangle: starts at active cell, sized to match source
4. Overwrite all cells in target rectangle
5. Record as undoable batch operation

**Rules:**
- Primary selection is ignored for determining target location—only active cell matters
- Paste always overwrites; never inserts/shifts cells
- If target extends beyond sheet bounds: operation fails with error message
- If clipboard contains single value: paste to single cell only
- If clipboard contains multi-cell data: paste entire rectangle

### Formula Adjustment on Paste

Formulas are adjusted based on the delta from original copy location:
- Delta = (paste_row - copy_row, paste_col - copy_col)
- Each relative reference shifts by delta
- Absolute references preserved

**Example:**
- Copy `=A1+B2` from C3
- Paste to E5 (delta: +2 rows, +2 cols)
- Result: `=C3+D4`

### While Editing (Edit Mode Paste)

**Behavior:**
1. Read from clipboard
2. Take first line only, trim whitespace
3. Insert at cursor position in edit buffer
4. If text selection exists: replace selection

---

## Fill Down Operation

### Ctrl+D / Data → Fill Down

**Precondition:** Primary selection must span at least 2 rows

**Source:** First row of primary selection
**Target:** All rows below first row in primary selection

**Behavior:**
1. For each column in selection:
   - Read source cell (first row)
   - For each target cell (rows below):
     - If source is formula: adjust references by row delta
     - If source is value: copy verbatim
     - Write to target cell
2. Record as undoable batch operation

**Rules:**
- Additional selections are ignored
- If only 1 row selected: show error, no action
- Formulas adjust row references, column references unchanged
- Selection is preserved after fill

**Example:**
```
Before (A1:A3 selected):
A1: =B1*2
A2: (empty)
A3: (empty)

After Ctrl+D:
A1: =B1*2
A2: =B2*2
A3: =B3*2
```

---

## Fill Right Operation

### Ctrl+R / Data → Fill Right

**Precondition:** Primary selection must span at least 2 columns

**Source:** First column of primary selection
**Target:** All columns after first column in primary selection

**Behavior:**
1. For each row in selection:
   - Read source cell (first column)
   - For each target cell (columns right):
     - If source is formula: adjust references by column delta
     - If source is value: copy verbatim
     - Write to target cell
2. Record as undoable batch operation

**Rules:**
- Additional selections are ignored
- If only 1 column selected: show error, no action
- Formulas adjust column references, row references unchanged
- Selection is preserved after fill

**Example:**
```
Before (A1:C1 selected):
A1: =A2+10
B1: (empty)
C1: (empty)

After Ctrl+R:
A1: =A2+10
B1: =B2+10
C1: =C2+10
```

---

## Fill Handle (Future - Not v1)

> **Note:** Fill handle is out of scope for v1. This section documents intended future behavior.

**Visual:** Small square at bottom-right corner of primary selection

**Behavior:**
- Drag down/right: extend source pattern into target range
- Pattern detection:
  - Single value: repeat
  - Numeric sequence: extend (1,2,3 → 4,5,6)
  - Date sequence: extend by detected interval
  - Formula: copy with reference adjustment
- Drag up/left: same logic, reversed direction

---

## Invariants (Always True)

1. **Copy/Cut operate on primary selection only**
2. **Paste targets active cell, not selection**
3. **Fill operates within primary selection; source is first row/column**
4. **Additional selections never participate in fill/copy/paste**
5. **All operations are undoable as single batch**
6. **Formulas always adjust relative references; absolute references never change**
7. **Paste always overwrites; never inserts or shifts**
8. **Operations that would exceed sheet bounds fail gracefully with error**
9. **Fill operations overwrite target cells using the same rules as Paste; no shifting or insertion occurs**

---

## Shape Matching Rules

### Paste Shape Matching

| Source Size | Target Start | Result |
|-------------|--------------|--------|
| 2×3 | Active cell D4 | Paste to D4:E6 |
| 1×1 | Active cell A1 | Paste to A1 only |
| 5×5 | Active at edge | Error if exceeds bounds |

- Target size always equals source size
- No tiling, no clipping, no smart resize
- Shape mismatch with selection: ignore selection, use source size

### Fill Shape

- Fill Down: source is 1 row × N columns, target is M rows × N columns
- Fill Right: source is N rows × 1 column, target is N rows × M columns
- Selection must be rectangular (guaranteed by selection semantics)

---

## Interaction with Additional Selections

**v1 Rule: Additional selections are ignored for all fill/copy/paste operations.**

Rationale:
- Reduces combinatorial complexity
- Avoids ambiguity about ordering and shape matching
- Additional selections primarily exist for formatting operations

Future versions may support:
- Copy each additional selection as separate clipboard item
- Paste to each additional selection (if shapes match)
- Multi-fill across discontinuous ranges

---

## Explicit Non-Goals (v1)

We do **not** support:

- **Paste Special** (paste values only, paste formats only, transpose)
- **Insert Paste** (shift cells down/right to make room)
- **Fill Handle** (drag to extend)
- **Series Fill** (smart pattern detection: 1,2,3 → 4,5,6)
- **Flash Fill** (AI pattern matching)
- **Fill across sheets**
- **Copy/paste across multiple selections**
- **Paste to non-matching shapes with tiling or clipping**
- **Cut with insert** (remove and shift cells)

These are explicitly deferred to future versions.

---

## Error Conditions

| Condition | Result |
|-----------|--------|
| Fill Down with <2 rows | Error message, no action |
| Fill Right with <2 columns | Error message, no action |
| Paste would exceed sheet bounds | Error message, no action |
| Empty clipboard on paste | No action, no error |
| Cut in edit mode | Cut selected text from edit buffer |

---

## State Transitions

```
Copy/Cut:
  Selection → Clipboard populated → Selection unchanged

Paste:
  Clipboard + Active Cell → Target cells overwritten → Selection unchanged

Fill Down/Right:
  Selection → Source replicated to target → Selection unchanged
```

All operations preserve selection state. The user's view of what's selected does not change.

---

## Undo Behavior

| Operation | Undo Unit | Restored State |
|-----------|-----------|----------------|
| Copy | N/A (no mutation) | N/A |
| Cut | Single batch | Original values restored |
| Paste | Single batch | Original values restored |
| Fill Down | Single batch | Original values restored |
| Fill Right | Single batch | Original values restored |

Undo restores cell values only; selection/clipboard are not affected.

---

## Version History

- **v1** (2026-01): Initial specification
