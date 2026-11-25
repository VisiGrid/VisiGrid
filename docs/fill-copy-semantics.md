# VisiGrid Fill & Copy Semantics (v2)

> **This document defines v2 behavior. Changes require a version bump and are treated as breaking API changes. Do not modify without explicit approval.**
>
> **If observed behavior differs from this document, the behavior is a bug.**

Fill & copy semantics define what happens when data is replicated, propagated, or inserted based on a selection. If selection semantics define *what is selected*, fill & copy semantics define *what happens to data when you act on that selection*.

---

## Definitions

### Source Range
- The cells providing data for an operation
- For Copy: the primary selection at time of copy
- For Fill Down: the first row of the primary selection; if the selection is a
  single row, the row above the selection
- For Fill Right: the first column of the primary selection; if the selection
  is a single column, the column to the left of the selection

### Target Range
- The cells receiving data from an operation
- An ordinary rectangular paste starts at the top-left of the primary selection and matches the source dimensions.
- For Fill Down: all rows below the first in primary selection; if the
  selection is a single row, the selected row itself
- For Fill Right: all columns after the first in primary selection; if the
  selection is a single column, the selected column itself

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


## Cut and fill with saved Table criteria

Cut, Fill Down/Right, Ctrl+Enter and the fill handle can operate in safe cells above/below a filtered Table or on another sheet. Resolve source and destination rows from the pre-edit projection, preserve canonical formula offsets, and skip hidden rows. A continuous cut/fill rectangle cannot cross the projected Table body boundary. Validate sources as well as targets; merged cells, headers, pivot output and spill receivers are refused. A failed cut leaves the clipboard unchanged. Each successful operation recalculates dependent Table views and uses one sparse guarded undo step; paste after cut remains a separate step. See [Table semantics](table-semantics.md#filling-visible-records) for limits and the existing Ctrl+Enter/series-fill extensions.

---

## Paste Operation

### Ctrl+V / Edit → Paste

**Target:** Top-left cell of the primary selection on an ordinary sheet.

**Behavior:**
1. Read from system clipboard (fall back to internal clipboard)
2. Parse as TSV (detect rows and columns)
3. Size the rectangle from that cell using the source row and column counts.
4. Overwrite all cells in target rectangle
5. Record as undoable batch operation

**Rules:**
- With a single cell selected, the active cell is the top-left.
- Paste always overwrites; never inserts/shifts cells
- If target extends beyond sheet bounds: operation fails with error message
- With one copied cell, Contents and All fill the selected cells and preserve the selection. Values, Formulas and Formats write and select only the top-left cell.
- After a successful multi-cell paste (Contents, All, Values, Formulas or Formats), select the written rectangle with its top-left active.

### Paste while Table filters are active

In a Table view, multi-cell paste starts at the active cell. [Table view paste rules](table-semantics.md#editing-and-pasting-visible-records)

Ordinary paste can target safe cells above/below a Table or on another sheet, as well as visible Table body records. It resolves every destination through the current row projection, validates all targets, recalculates a candidate workbook and checks every saved Table view before publication. A paste beginning in a Table body cannot overflow that body. Header/adjacent-body/protected targets, invalid bounds, merged clipboard sources and unsafe recalculation reject the whole operation. One sparse undo step restores the edit and recomputes dependent filters. See [Table semantics](table-semantics.md#editing-outside-filtered-tables).

### Table header paste

A paste starting on a Table header is an atomic schema rename. Paste/Paste Values accepts one row contained in the same Table, preserves column IDs and header formatting/comments, and rewrites dependent formulas. The selection must contain only that Table's headers, with no additional selections. Blank/duplicate/invalid names, multirow data and overflow reject the whole operation. Single-name broadcasting is refused. Saved Table filters/sorts remain bound to their columns. One undo reverses all renamed headers and dependent rewrites. See [Table semantics](table-semantics.md#multi-header-schema-paste) for clipboard and replay details.

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

**Source:** First row of primary selection. If the selection spans only 1 row,
the row above the selection (Excel-style single-cell/single-row fill).
**Target:** All rows below the source row, within the selection.

**Precondition:** Multi-row selection, or single-row selection not on row 1.

**Behavior:**
1. For each column in selection:
   - Read source cell (source row)
   - For each target cell (rows below source, within selection):
     - If source is formula: adjust references by row delta
     - If source is value: copy verbatim
     - Write to target cell
2. Record as undoable batch operation

**Rules:**
- Additional selections are ignored
- If only 1 row selected: fill from the row above into the selection
- If only 1 row selected AND it is row 1: show error, no action
- Formulas adjust row references, column references unchanged
- Selection is preserved after fill

**Example (multi-row):**
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

**Example (single cell):**
```
Before (A2 selected):
A1: =B1*2
A2: (empty)

After Ctrl+D:
A1: =B1*2
A2: =B2*2
```

---

## Fill Right Operation

### Ctrl+R / Data → Fill Right

**Source:** First column of primary selection. If the selection spans only 1
column, the column to the left of the selection (Excel-style
single-cell/single-column fill).
**Target:** All columns after the source column, within the selection.

**Precondition:** Multi-column selection, or single-column selection not on
column A.

**Behavior:**
1. For each row in selection:
   - Read source cell (source column)
   - For each target cell (columns right of source, within selection):
     - If source is formula: adjust references by column delta
     - If source is value: copy verbatim
     - Write to target cell
2. Record as undoable batch operation

**Rules:**
- Additional selections are ignored
- If only 1 column selected: fill from the column to the left into the selection
- If only 1 column selected AND it is column A: show error, no action
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
2. **An ordinary rectangular paste starts at the primary selection’s top-left.**
3. **Fill operates within primary selection; source is first row/column**
4. **Fill, copy and rectangular paste use only the primary selection.**
5. **All operations are undoable as single batch**
6. **Formulas always adjust relative references; absolute references never change**
7. **Paste always overwrites; never inserts or shifts**
8. **Operations that would exceed sheet bounds fail gracefully with error**
9. **Fill operations overwrite target cells using the same rules as Paste; no shifting or insertion occurs**

---

## Shape Matching Rules

### Paste Shape Matching

| Source (rows × columns) | Primary selection | Result |
|-------------|--------------|--------|
| 2×3 | D4 | Paste to D4:F5 |
| 1×1 | A1 | Paste to A1 only |
| 5×5 | Top-left at sheet edge | Error if exceeds bounds |
| 5×1 | F7 dragged up to F3 (active F7) | Paste to F3:F7 with F3 active |

- A rectangular paste uses the source dimensions even when the selection has a different shape.
- No tiling, no clipping, no smart resize

### Fill Shape

- Fill Down: source is 1 row × N columns, target is M rows × N columns
- Fill Right: source is N rows × 1 column, target is N rows × M columns
- Selection must be rectangular (guaranteed by selection semantics)

---

## Interaction with Additional Selections

Rectangular paste skips additional Ctrl+click selections. A one-cell Contents or All broadcast fills them along with the primary selection.

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

The original v1 scope excluded the following operations:

- **Paste Special** (paste values only, paste formats only, transpose)
- **Insert Paste** (shift cells down/right to make room)
- **Fill Handle** (drag to extend)
- **Series Fill** (smart pattern detection: 1,2,3 → 4,5,6)
- **Flash Fill** (AI pattern matching)
- **Fill across sheets**
- **Copy/paste across multiple selections**
- **Paste to non-matching shapes with tiling or clipping**
- **Cut with insert** (remove and shift cells)

The v2 paste rules above cover Contents, All, Values, Formulas and Formats.

---

## Error Conditions

| Condition | Result |
|-----------|--------|
| Fill Down with 1 row selected on row 1 | Error message, no action |
| Fill Right with 1 column selected on column A | Error message, no action |
| Paste would exceed sheet bounds | Error message, no action |
| Empty clipboard on paste | No action, no error |
| Cut in edit mode | Cut selected text from edit buffer |

---

## State Transitions

```
Copy/Cut:
  Selection → Clipboard populated → Selection unchanged

Paste:
  Source rectangle + primary selection → Written rectangle selected, top-left active
  One-cell Contents or All broadcast → Selected cells filled → Selection preserved

Fill Down/Right:
  Selection → Source replicated to target → Selection unchanged
```

Copy, cut, fill and one-cell broadcasts preserve selection.

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

- **v2** (2026-11): Ordinary rectangular paste starts at the primary selection’s top-left and selects the written rectangle.
- **v1** (2026-01): Initial specification
