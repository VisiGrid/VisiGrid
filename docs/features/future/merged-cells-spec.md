# Merged Cells Specification

> **Merged cells combine multiple cells into one.** Create headers, labels, and layouts that span rows and columns.

---

## Purpose

Merged cells exist for **visual layout** — creating headers, labels, and structured layouts that don't fit the single-cell model.

**Use cases:**
- Report headers spanning multiple columns
- Category labels spanning rows
- Centered titles above data tables
- Invoice/form layouts
- Calendar and schedule grids

---

## Behavior

### Merging Cells

**Entry points:**
- Format menu → Merge Cells → Merge & Center
- Command palette: "Merge Cells"
- Context menu: Merge Cells
- Keyboard: Ctrl+Shift+M (see Keyboard Shortcuts for rationale)

**Merge options:**
| Option | Result |
|--------|--------|
| Merge & Center | Merge all cells into one region, center content horizontally |
| Merge Across | Merge each row separately (see below) |
| Merge Cells | Merge all cells into one region, keep original alignment |
| Unmerge Cells | Split merged cell back to individual cells |

**Merge Across** creates N separate `MergedRegion`s, one per row in the selection:

```
Selection: A1:C3 → Merge Across produces:
  MergedRegion { start: (0,0), end: (0,2) }  // A1:C1
  MergedRegion { start: (1,0), end: (1,2) }  // A2:C2
  MergedRegion { start: (2,0), end: (2,2) }  // A3:C3
```

Data loss warning for Merge Across lists affected cells per row: "Data in B1, C1 (row 1), B2, C2 (row 2), ..." — or an aggregate count if the list is long.

### Data Handling

**When merging cells with data:**
- Only top-left cell's value is preserved
- Other cells' values are discarded (with warning)
- Formulas referencing merged cells reference the merged cell

**Warning dialog (if data would be lost):**
```
┌─────────────────────────────────────────────┐
│ Merge Cells                              ✕  │
├─────────────────────────────────────────────┤
│ ⚠ Merging cells only keeps the upper-left  │
│   value and discards other values.          │
│                                             │
│   Data in B2, B3, C2, C3 will be lost.      │
│                                             │
│              [Cancel]  [Merge Anyway]       │
└─────────────────────────────────────────────┘
```

### Unmerging Cells

- Split merged cell back into original cells
- Value stays in top-left cell only
- Other cells become empty
- All unmerged cells **inherit the origin cell's format** (bold, borders, background, etc.)
- This matches Excel behavior and avoids jarring format loss

---

## Visual Representation

### Merged Cell Display

```
Before merge:
┌───────┬───────┬───────┐
│   A   │   B   │   C   │
├───────┼───────┼───────┤
│ Title │       │       │  ← A1, B1, C1 separate
├───────┼───────┼───────┤
│ Data  │ 123   │ 456   │
└───────┴───────┴───────┘

After merge A1:C1:
┌───────────────────────┐
│        Title          │  ← Single merged cell
├───────┬───────┬───────┤
│ Data  │ 123   │ 456   │
└───────┴───────┴───────┘
```

### Border Rendering

Borders use **outer-edge composition with per-side edge sampling**:

- When rendering a cell edge, sample the adjacent cell on the other side of the edge
- If both cells belong to the same merge → **suppress the border** (internal edge)
- If the adjacent cell is outside the merge (or a different merge) → **draw the border** (outer edge)

**Border style is NOT read from the origin cell alone.** Each outer edge is resolved from the cells on that edge using max-precedence:

- **Top edge:** `border_top` from each cell in the top row of the merge (row `start.0`, cols `start.1..=end.1`)
- **Bottom edge:** `border_bottom` from each cell in the bottom row
- **Left edge:** `border_left` from each cell in the left column
- **Right edge:** `border_right` from each cell in the right column

**Precedence order:** None < Hair < Thin < Medium < Thick (dashed/dotted treated as same weight as their solid equivalent).

This matters because Excel files frequently store different border styles on different edge cells of a merge. Reading only the origin would lose the right and bottom borders entirely.

**Performance note:** For wide merges (e.g., 1 row × 16k columns), sampling every edge cell per frame is wasteful. Cache resolved borders per merge:

```rust
struct ResolvedMergeBorders {
    top: CellBorder,
    right: CellBorder,
    bottom: CellBorder,
    left: CellBorder,
}
```

Invalidate when any `style_id` in the merge changes or the merge is resized. This is optional for initial implementation but will matter when scrolling large financial models.

```
Example: A1:C2 merged

Top border: max(A1.border_top, B1.border_top, C1.border_top)
Right border: max(C1.border_right, C2.border_right)
Bottom border: max(A2.border_bottom, B2.border_bottom, C2.border_bottom)
Left border: max(A1.border_left, A2.border_left)

Edge between A1 and B1: both in same merge → no border drawn
Edge between C1 and D1: C1 in merge, D1 outside → draw right border
Edge between A2 and A3: A2 in merge, A3 outside → draw bottom border
```

### Selection

- Click anywhere in merged cell → entire merged region selected
- Selection highlight covers entire merged area
- Active cell indicator shows top-left position
- **Selection clamping:** If a selection partially overlaps a merge, expand the selection to include the entire merge region (Excel behavior)

---

## Navigation

### Arrow Keys

| Current Cell | Key | Result |
|--------------|-----|--------|
| Before merged cell | → | Jump to merged cell |
| In merged cell | → | Jump past merged cell |
| Above merged cell | ↓ | Jump to merged cell |
| In merged cell | ↓ | Jump past merged cell |

**Rule:** Merged cell acts as single cell for navigation. Arrow keys skip over the merged region.

### Tab/Enter

- Tab from merged cell → next cell after merged region
- Enter from merged cell → cell below merged region
- Shift+Tab → cell before merged region
- Shift+Enter → cell above merged region

### Ctrl+Arrow

- Ctrl+Arrow jumps to edge of data region
- Merged cells count as single data point
- Stops at merged cell boundary

### Go To

- Go To "B2" when B2 is part of merge → selects the merged cell
- Reference shown as "A1:C1" (the merge range)

---

## Editing

### Enter Edit Mode

- Double-click anywhere in merged cell → edit mode
- F2 → edit mode
- Start typing → edit mode
- Content appears in formula bar

### Edit Display

- Cursor and text editing happens in merged cell area
- Text wraps within merged cell bounds (visual only — see note below)
- Alignment applied to merged cell as a whole

**v1 wrapping limitation:** Text wrapping is visual-only; there is no auto row-height expansion for merged cells. Long text will visually clip at the merge bounds. Wrapping only applies if `wrapText` is set in the cell format (matching Excel). Auto row-height for merged cells is a future enhancement.

---

## Formulas

### Referencing Merged Cells

**Single-cell references** redirect to the merge origin:

```
Merged: A1:C1 contains "Header"

=A1        → "Header"
=B1        → "Header" (redirects to origin A1)
=C1        → "Header" (redirects to origin A1)
```

**Range references** still iterate all cells — hidden cells in a merge are treated as empty:

```
Merged: A1:C1 contains "Header"

=COUNTA(A1:C1)  → 1  (only origin has value, B1/C1 are empty)
=A1:C1          → "Header" (single-value context returns origin)
```

This distinction matters for SUM, COUNTA, and other aggregate functions — they should not double-count merged values.

**Error behavior:**
- Single-cell redirect returns exactly the origin's `CellValue`, including errors (`#DIV/0!`, `#REF!`, etc.)
- Hidden cells in range iteration behave as **truly empty** (not zero, not error) — `SUM` skips them, `COUNTA` doesn't count them, `AVERAGE` excludes them from both numerator and denominator

**Dynamic range functions:** Functions that construct ranges at eval time (`OFFSET`, `INDIRECT` returning a range) follow the same rule — hidden merged cells are empty. No special casing is needed. This matches Excel.

### Formulas IN Merged Cells

- Formula entered in merged cell works normally
- Computed value displayed across merged region
- References adjust normally if merged cell is moved

### INDIRECT/OFFSET to Merged Cells

- `=INDIRECT("B1")` where B1 is merged → returns merged cell value
- Dynamic references resolve to the merge

---

## Data Model

### MergedRegion

```rust
pub struct MergedRegion {
    pub start: (usize, usize),  // Top-left (row, col)
    pub end: (usize, usize),    // Bottom-right (row, col)
}

impl MergedRegion {
    pub fn rows(&self) -> usize {
        self.end.0 - self.start.0 + 1
    }

    pub fn cols(&self) -> usize {
        self.end.1 - self.start.1 + 1
    }

    pub fn contains(&self, row: usize, col: usize) -> bool {
        row >= self.start.0 && row <= self.end.0 &&
        col >= self.start.1 && col <= self.end.1
    }

    pub fn top_left(&self) -> (usize, usize) {
        self.start
    }
}
```

### Storage

The naive `Vec<MergedRegion>` with linear scan is O(n) per lookup. Since every cell render, navigation step, and formula evaluation calls `get_merge()`, this must be fast.

**Primary storage:** `Vec<MergedRegion>` (canonical list, used for serialization and iteration).

**Lookup index:** `HashMap<(usize, usize), usize>` mapping every `(row, col)` in any merge to its index in the `Vec`. This gives O(1) lookup at the cost of rebuilding the map when merges change.

```rust
pub struct Sheet {
    // ... existing fields ...
    pub merged_regions: Vec<MergedRegion>,

    /// (row, col) → index into merged_regions. Covers ALL cells in every merge.
    /// Rebuilt by rebuild_merge_index() after any merge mutation.
    #[serde(skip)]
    merge_index: HashMap<(usize, usize), usize>,
}

impl Sheet {
    /// Rebuild the merge lookup index. Call after any merge mutation.
    fn rebuild_merge_index(&mut self) {
        self.merge_index.clear();
        for (idx, m) in self.merged_regions.iter().enumerate() {
            for r in m.start.0..=m.end.0 {
                for c in m.start.1..=m.end.1 {
                    self.merge_index.insert((r, c), idx);
                }
            }
        }
    }

    /// Find merged region containing cell, if any. O(1).
    pub fn get_merge(&self, row: usize, col: usize) -> Option<&MergedRegion> {
        self.merge_index.get(&(row, col)).map(|&idx| &self.merged_regions[idx])
    }

    /// Check if cell is the top-left of a merge
    pub fn is_merge_origin(&self, row: usize, col: usize) -> bool {
        self.get_merge(row, col).map_or(false, |m| m.start == (row, col))
    }

    /// Check if cell is part of a merge but not the origin
    pub fn is_merge_hidden(&self, row: usize, col: usize) -> bool {
        self.get_merge(row, col).map_or(false, |m| m.start != (row, col))
    }
}
```

**Tradeoff:** For a sheet with M merges averaging S cells each, the index uses O(M*S) memory. For typical spreadsheets (< 1000 merges, < 100 cells each) this is negligible.

**Large-merge escape hatch:** If any single `MergedRegion` exceeds 50,000 cells (e.g., a full-row merge across 16k columns), the HashMap approach becomes wasteful. At that threshold, switch the sheet's index to a row-interval representation (`BTreeMap<usize, Vec<(col_start, col_end, region_idx)>>`) that stores one entry per merge per row instead of one entry per cell. This is a future optimization — implement the HashMap first, add the interval index only if real-world files hit the threshold.

The index is `#[serde(skip)]` and rebuilt on deserialization / after any merge mutation (add, remove, insert/delete rows/cols).

### Cell Value Storage

Only the top-left (origin) cell stores the value. Both reads and writes redirect to the origin:

```rust
impl Sheet {
    pub fn get_value(&self, row: usize, col: usize) -> &Value {
        // If cell is in a merge, return top-left cell's value
        if let Some(merge) = self.get_merge(row, col) {
            let (origin_row, origin_col) = merge.top_left();
            return self.get_cell(origin_row, origin_col).value();
        }
        self.get_cell(row, col).value()
    }

    pub fn set_value(&mut self, row: usize, col: usize, value: &str) {
        // If cell is in a merge, redirect write to origin
        let (target_row, target_col) = if let Some(merge) = self.get_merge(row, col) {
            merge.top_left()
        } else {
            (row, col)
        };
        // ... existing set_value logic on (target_row, target_col) ...
    }
}
```

This prevents data from being silently written to hidden cells inside a merge.

### Style Storage for Hidden Cells

Hidden cells inside a merge retain their `style_id` and `format` as imported — they are not stripped. Only **value/formula** is exclusive to the origin cell.

Rationale:
- XLSX files frequently store border/fill styles on individual cells within a merge (not just the origin)
- Keeping per-cell styles preserves Excel fidelity for border composition (see Border Rendering)
- On unmerge, each cell's existing style is already correct — no reconstruction needed
- This aligns with the existing `style_id` + `style_table` architecture from formatting import

---

## Operations on Merged Cells

### Copy/Paste

**Copy merged cell:**
- Clipboard contains merged region information + merge shape
- Paste recreates the merge at destination (if space allows)

**Paste rules (strict — block ambiguous cases in v1):**

| Scenario | Result |
|----------|--------|
| Single cell → any cell inside a merge | Paste value into merge origin (regardless of which cell in the merge is active) |
| Range → destination with no merges | Normal paste |
| Range → destination partially overlapping a merge | **Block:** "Cannot paste — would split merged cells. Unmerge first." |
| Clipboard contains merge → destination has enough free area | Recreate merge at destination |
| Clipboard contains merge → destination overlaps existing merge | **Block:** "Cannot paste — would overlap existing merged cells." |
| Clipboard shape doesn't match destination merge shape | **Block:** "Cannot paste — shape mismatch with merged cells." |

**Rationale:** Excel is strict about paste + merged cells. Being lenient here creates silent data corruption. Better to block and let the user unmerge first than to guess.

### Fill

**Fill Down/Right with merged cells:**
- Merged cells in source → replicate merge structure
- Fill into merged cells → fills the top-left cell only

### Delete/Clear

- Delete content in merged cell → clears the top-left cell value
- Does NOT unmerge
- Clear All → clears value AND unmerges (optional behavior)

### Insert Rows/Columns

Insertion is defined in terms of **grid-line indices**. Inserting a row "at index k" means inserting between row k-1 and row k (i.e., the new row occupies index k, pushing existing row k downward).

For a merge spanning rows `start_row..=end_row`:

| Insert position k | Effect |
|--------------------|--------|
| k <= start_row | **Shift:** merge moves down (start_row += n, end_row += n) |
| start_row < k <= end_row | **Expand:** merge grows (end_row += n) |
| k > end_row | No effect |

Same logic applies to column insertion with `start_col`/`end_col`.

```
Example: Merge A2:C4 (rows 1..=3, cols 0..=2)

Insert 1 row at k=0 (above merge):  → A3:C5 (shift)
Insert 1 row at k=1 (at merge top):  → A3:C5 (shift — boundary = shift, not expand)
Insert 1 row at k=2 (inside merge): → A2:C5 (expand)
Insert 1 row at k=4 (at merge bottom+1): → A2:C4 (no effect)
```

**Key rule:** Insert at `start_row` shifts the merge (the insertion is *above* the merge). Insert at `start_row + 1` through `end_row` expands it.

### Delete Rows/Columns

For a merge spanning rows `start_row..=end_row`, deleting row k:

| Delete position k | Effect |
|--------------------|--------|
| k < start_row | **Shift:** merge moves up (start_row -= 1, end_row -= 1) |
| start_row <= k <= end_row | **Shrink:** end_row -= 1 |
| k > end_row | No effect |

If shrinking causes `start_row > end_row` or `start_col > end_col`, the merge is fully deleted (handled by `normalize_merges()`).

Same logic applies to column deletion with `start_col`/`end_col`.

### Normalize After Insert/Delete

After any row/column insert or delete, call `normalize_merges()` which:

1. Adjusts all merge coordinates for the shift
2. Clamps merges that extend beyond sheet bounds
3. Removes **degenerate merges** that collapsed to 1×1 (single cell)
4. Rebuilds the merge index

```rust
impl Sheet {
    pub fn normalize_merges(&mut self) {
        // Remove degenerate merges (start == end after shrink)
        self.merged_regions.retain(|m| m.rows() > 1 || m.cols() > 1);
        self.rebuild_merge_index();
        self.debug_assert_no_merge_overlap();
    }

    /// Debug-only invariant: no two merges may cover the same cell.
    /// Called after normalize_merges(), XLSX import, and insert/delete operations.
    #[cfg(debug_assertions)]
    fn debug_assert_no_merge_overlap(&self) {
        let mut seen = HashSet::new();
        for m in &self.merged_regions {
            for r in m.start.0..=m.end.0 {
                for c in m.start.1..=m.end.1 {
                    assert!(
                        seen.insert((r, c)),
                        "Overlapping merged cells at ({}, {})", r, c
                    );
                }
            }
        }
    }

    #[cfg(not(debug_assertions))]
    fn debug_assert_no_merge_overlap(&self) {}
}
```

This is called as a post-step after the existing insert/delete row/column logic, keeping the merge adjustment separate from the row/column mechanics. The debug assertion catches silent corruption from XLSX import edge cases or normalize logic bugs.

### Sort

**Sort range containing merged cells:**
- Block with error: "Cannot sort a range containing merged cells. Unmerge cells first."
- No options, no partial workarounds — just a clear message and bail

### Filter

**Filter column with merged cells:**
- Block with error: "Cannot filter a range containing merged cells. Unmerge cells first."
- Same approach as sort — merged cells and row-level visibility are fundamentally incompatible

---

## Rendering

### Grid Rendering

```rust
fn render_cell(&self, row: usize, col: usize, cx: &mut Context) {
    // Skip rendering if cell is hidden part of merge
    if self.sheet().is_merge_hidden(row, col) {
        return; // Don't render this cell
    }

    // If this is merge origin, render expanded
    if let Some(merge) = self.sheet().get_merge(row, col) {
        if merge.start == (row, col) {
            self.render_merged_cell(merge, cx);
            return;
        }
    }

    // Normal cell rendering
    self.render_normal_cell(row, col, cx);
}

fn render_merged_cell(&self, merge: &MergedRegion, cx: &mut Context) {
    // Calculate bounds spanning all cells in merge
    let bounds = self.calculate_merge_bounds(merge);

    // Render single cell spanning the region
    // Background, borders, content
}
```

### Selection Rendering

```rust
fn render_selection(&self, cx: &mut Context) {
    let (row, col) = self.selected;

    // If selected cell is in merge, highlight entire merge
    if let Some(merge) = self.sheet().get_merge(row, col) {
        let bounds = self.calculate_merge_bounds(merge);
        self.render_selection_highlight(bounds, cx);
    } else {
        // Normal selection
    }
}
```

### Merge-Aware Layout Requirements

**Hit-testing:** Mouse clicks must resolve to the correct merge region. When a click lands on a hidden cell inside a merge, map it to the merge's origin cell.

**CenterAcrossSelection precedence:** Merged cells take priority over CenterAcrossSelection alignment. If a cell is part of a merge, its merge region defines the visual span — the CenterAcross scan should not cross merge boundaries.

**Frozen panes:** A merge that spans the frozen/unfrozen boundary is an edge case. For v1, clip the merge rendering at the freeze boundary.

Behavior in clipped state:
- Click on either visible portion of the clipped merge → selects the full merge (logical selection is the whole region)
- Selection rectangle is clipped in rendering but remains logically complete
- Arrow key navigation skips based on the full region, not the visible portion
- If this proves confusing in practice, a future version can block merges from crossing freeze boundaries

**Excel compatibility note:** Excel allows merges across freeze panes but its rendering behavior is inconsistent (sometimes visually breaks the merge). VisiGrid clips rendering but preserves logical selection. This is a known divergence — our behavior is more predictable but may differ from Excel in edge cases.

---

## Validation Rules

### Cannot Merge

- Selection includes part of existing merge (but not all)
- Selection is non-contiguous (additional selections)
- Selection spans hidden rows/columns (filtered)

### Merge Conflicts

```
Existing merge: A1:B2
New selection: B1:C2

Result: Error - "Selection overlaps existing merged cells"
Options: Unmerge first, or extend selection
```

---

## File Formats

### .vgrid (Native)

```sql
CREATE TABLE merged_regions (
    sheet_id INTEGER,
    start_row INTEGER,
    start_col INTEGER,
    end_row INTEGER,
    end_col INTEGER,
    PRIMARY KEY (sheet_id, start_row, start_col)
);
```

### XLSX

**Import:**
- Read `<mergeCells>` from sheet XML
- Create MergedRegion for each `<mergeCell ref="A1:C3"/>`
- **Border composition:** Excel stores borders on individual cells, not on the merge region. On import, keep per-cell styles as-is (see Style Storage for Hidden Cells). Border composition for rendering is computed at render time from edge cells, not flattened at import.
- **Overlap policy:** If the XLSX contains overlapping merge regions (malformed files — it happens), drop the later merge and record `"invalid merge overlap at A1:C3 (conflicts with A1:B2)"` in the import report's `unsupported_format_features`. Do not panic or corrupt the merge index.

**Export:**
- Write `<mergeCells>` element
- One `<mergeCell>` per MergedRegion
- Since hidden cells retain their `style_id` (see Style Storage), their per-cell border properties are already available. Write each cell's own border to the XLSX output. If a cell inside a merge has no style, fall back to the origin cell's borders for that edge.

### CSV

- Export: Merged cell value in top-left position, empty in others
- Import: No merge information (plain text)

---

## Keyboard Shortcuts

| Action | Windows/Linux | macOS |
|--------|---------------|-------|
| Merge & Center | Ctrl+Shift+M | Cmd+Shift+M |
| Unmerge | Ctrl+Shift+U | Cmd+Shift+U |

**Why not Ctrl/Cmd+M?** On macOS, Cmd+M is the system shortcut for "Minimize window." Overriding it would break platform conventions. Using Ctrl+Shift+M on both platforms keeps shortcuts consistent cross-platform. Excel uses Alt+H,M,C (ribbon path) — there is no universal standard shortcut for merge.

---

## Menu Integration

### Format Menu

```
Format
├── ...
├── Merge Cells                        ▸
│   ├── Merge & Center     Ctrl+Shift+M
│   ├── Merge Across
│   ├── Merge Cells
│   └── Unmerge Cells      Ctrl+Shift+U
├── ...
```

### Context Menu

```
...
├── Merge Cells            Ctrl+Shift+M  │  ← If selection can merge
├── Unmerge Cells          Ctrl+Shift+U  │  ← If selection is merged
...
```

---

## Inspector

When merged cell selected:

```
┌─────────────────────────────────┐
│ Merged Cell A1:C3               │
├─────────────────────────────────┤
│ Value: "Header"                 │
│ Type: Text                      │
│ Span: 3 columns × 3 rows        │
├─────────────────────────────────┤
│ [Unmerge]                       │
└─────────────────────────────────┘
```

---

## Tests (Required)

### Basic Merge/Unmerge
1. Select A1:C1 → Merge → single merged cell
2. Click merged cell → entire region selected
3. Unmerge → three separate cells, value in A1 only
4. Merge with data in multiple cells → warning, only A1 kept

### Navigation
5. Arrow right into merge → lands on merge
6. Arrow right from merge → skips to cell after merge
7. Tab from merge → next cell after merge
8. Go To "B1" when B1 is in merge → selects merge

### Editing
9. Double-click merge → edit mode
10. Type in merge → value stored in top-left
11. Formula bar shows merged cell reference

### Formulas
12. =A1 where A1 is merged → returns value
13. =B1 where B1 is part of A1:C1 merge → returns A1 value
14. Formula in merged cell → computes correctly

### Operations
15. Copy merged cell → paste recreates merge
16. Delete content → clears value, keeps merge
17. Insert row in merge → merge expands
18. Delete row in merge → merge shrinks

### Rendering
19. Merged cell renders as single cell
20. Borders around entire merge
21. Content centered/aligned correctly
22. Hidden cells not rendered

### Edge Cases
23. Merge entire row → works
24. Merge entire column → works (with limits)
25. Overlapping merge attempt → error
26. Sort with merged cells → blocked with error message
27. Filter with merged cells → blocked with error message
28. set_value on hidden cell → redirects to origin
29. Unmerge → all cells inherit origin format
30. Insert row at merge start → merge shifts down (not expand)
31. Insert row inside merge (start+1..end) → merge expands
32. Delete rows that collapse merge to 1×1 → normalize removes degenerate merge
33. Selection partially overlapping merge → expands to include full merge

### Paste + Merge Interactions
34. Paste into partially-overlapping merge → blocked with error
35. Paste clipboard with merge → recreated at destination
36. Paste clipboard with merge → blocked if destination has existing merge overlap

### Formula Edge Cases
37. =B1 where B1 is hidden in merge, origin has #DIV/0! → returns #DIV/0!
38. =COUNTA(A1:C1) with merge A1:C1 → 1 (hidden cells are empty, not zero)
39. =SUM(A1:C3) with merge A1:C3 containing 10 → 10 (not 10*9)

### Merge Across
40. Merge Across A1:C3 → creates 3 separate row merges
41. Merge Across with data in non-origin cells → warning lists per-row losses

### File Formats
42. Save/load .vgrid with merges → preserved (merge_index rebuilt on load)
43. Import XLSX with merges → merges created, per-cell styles preserved
44. Import XLSX with overlapping merges → later merge dropped, reported in import stats
45. Export XLSX with merges → merges in file, borders decomposed to edge cells
46. Export CSV → value in top-left only

---

## Not In Scope (v1)

| Feature | Rationale |
|---------|-----------|
| Unmerge and fill | Could fill value to all unmerged cells |
| Merge across sheets | Complexity |
| Partial merge overlap resolution | Keep simple: error on overlap |
| Merge in filtered view | Complexity |

---

## Complications & Tradeoffs

### Performance

- Every cell render, navigation, and formula eval calls `get_merge()` — must be O(1)
- The `merge_index: HashMap<(usize, usize), usize>` provides O(1) lookup (see Storage section)
- Index is rebuilt on merge mutation, not on every access
- For extremely large merges (entire-row spans), consider a row-interval index as a future optimization

### Clipboard Complexity

Different apps handle merged cell paste differently:
- Excel: Recreates merge
- Sheets: Sometimes recreates, sometimes doesn't
- VisiGrid: Recreate merge on paste

### Data Integrity

Merged cells can hide data:
- User merges over existing values
- Data appears "lost" (actually discarded)
- Important to warn clearly

---

## Implementation Order

**Phase 1: Data Model** — COMPLETE
- [x] MergedRegion struct (`sheet.rs:150`)
- [x] Storage in Sheet (`merged_regions: Vec<MergedRegion>`, `merge_index: HashMap`)
- [x] merge/unmerge operations (`add_merge()`, `remove_merge()`, overlap detection)
- [x] get_merge() O(1) lookup via HashMap index
- [x] `merge_origin_coord()` helper for value write redirect
- [x] `normalize_merges()` + `debug_assert_no_merge_overlap()`
- [x] `is_merge_origin()` / `is_merge_hidden()` predicates
- [x] `rebuild_merge_index()` called after all merge mutations
- [x] Serde roundtrip (merge_index is `#[serde(skip)]`, rebuilt on deser)

**Phase 2: Rendering** — COMPLETE
- [x] Skip hidden cells in render (spacer elements for hidden cells)
- [x] Render merged cell with expanded bounds (overlay elements)
- [x] Selection highlight for merged cells
- [x] Border composition (per-side edge sampling with precedence)
- [x] Hit-testing (click on hidden cell → redirect to origin)
- [x] Spill exclusion (spill ranges avoid merged regions)
- [ ] CenterAcrossSelection precedence over merge boundaries
- [ ] Frozen pane clipping for cross-boundary merges

**Phase 3: Navigation** — COMPLETE
- [x] Arrow key navigation over merges (skip merged region, `move_selection()`)
- [x] Tab/Enter navigation (next cell after merged region, `confirm_edit_and_move()`)
- [x] Ctrl+Arrow jump navigation (`jump_selection()`, `extend_jump_selection()`)
- [x] `find_data_boundary()` treats merged cells as occupied data units
- [x] Go To merged cells (B1 in merge → selects merge, `confirm_goto()`)
- [x] Tab-chain Enter/Shift+Enter snap to merge origin (`confirm_edit_enter()`, `confirm_edit_up_enter()`)
- [x] Selection extension over merges (`extend_selection()`)

**Phase 4: Operations** — PARTIALLY COMPLETE
- [ ] Copy/paste with merge recreation (clipboard merge metadata)
- [x] Insert/delete rows with merge adjustment (shift/expand/shrink/degenerate)
- [x] Insert/delete columns with merge adjustment
- [x] Clear content in merged cells (redirects to origin via `clear_cell()`)
- [x] Value writes redirect to origin (`set_value()`, `set_cycle_error()`)
- [x] Format writes do NOT redirect (per-cell styles preserved on hidden cells)
- [x] Paste blocked when partially overlapping merge (`paste_would_split_merge()`)
- [x] Sort blocked when merges exist (`block_if_merged`)
- [x] AutoFilter blocked when merges exist (`block_if_merged`)
- [x] Fill down/right blocked when merges exist (`block_if_merged`)
- [x] Fill handle (vertical/horizontal) blocked when merges exist (`block_if_merged`)
- [x] Trim whitespace blocked when merges exist (`block_if_merged`)
- [x] Replace All blocked when merges exist (`block_if_merged`)
- [x] Fill selection from primary blocked when merges exist (`block_if_merged`)
- [x] Formula CellRef redirect (single-cell refs → origin; ranges → no redirect)
- [x] Cross-sheet formula CellRef redirect (`get_merge_start_sheet`)

**Phase 5: UI** — NOT STARTED
- [ ] Format menu entries (Merge & Center, Merge Across, Merge Cells, Unmerge)
- [ ] Context menu entries
- [ ] Keyboard shortcuts (Ctrl+Shift+M, Ctrl+Shift+U)
- [ ] Merge warning dialog (data loss confirmation)
- [ ] Merge Across (per-row merge creation)

**Phase 6: File Formats** — PARTIALLY COMPLETE
- [x] .vgrid persistence (MergedRegion serialized via serde, merge_index rebuilt on load)
- [x] XLSX import (reads `<mergeCells>`, creates MergedRegion per `<mergeCell>`)
- [ ] XLSX export (write `<mergeCells>` element)
- [ ] CSV export (value in top-left only, empty in others)

**Phase 7: Edge Cases** — COMPLETE
- [x] Sort/filter blocked with clear error message
- [x] Overlap detection (`add_merge()` rejects overlapping merges)
- [x] Formula references to merged cells (CellRef redirect + range semantics)
- [x] `debug_assert_no_merge_overlap()` invariant check

---

## Implementation Status

**Phase 1 (Data Model): COMPLETE.** Full merge data model with O(1) lookup, add/remove/normalize operations, insert/delete row/column adjustment, serde roundtrip, and overlap detection.

**Phase 2 (Rendering): COMPLETE.** Overlay rendering for merged cells, spacer elements for hidden cells, spill exclusion, border composition, hit-testing, selection highlight. 17 rendering tests.

**Phase 3 (Navigation): COMPLETE.** All navigation functions are merge-aware: `move_selection()` and `extend_selection()` move from effective merge edges and snap/expand on landing. `jump_selection()` and `extend_jump_selection()` start from merge edges, snap to origin on landing, and expand selection for merges. `find_data_boundary()` treats merged cells as occupied data units for Ctrl+Arrow. `confirm_goto()` redirects to merge origin and selects the full merge. Tab-chain Enter/Shift+Enter snap to merge origin.

**Phase 4 (Operations — Semantic Correctness): MOSTLY COMPLETE.** All value write entry points redirect to merge origin. Formula CellRef evaluation redirects single-cell refs to origin while range iteration treats hidden cells as empty. 12 bulk operations are guarded with `block_if_merged()` or `paste_would_split_merge()`. Remaining: copy/paste with merge recreation.

**Phase 6 (File Formats): PARTIALLY COMPLETE.** .vgrid and XLSX import work. XLSX export and CSV export remain.

**Phase 7 (Edge Cases): COMPLETE.** Sort/filter blocked, overlap detection, formula redirect semantics all implemented and tested.

**Phase 5: NOT STARTED.** UI (menus/shortcuts/dialogs) is the remaining work to expose merges to users.

### What's Left Before Merge UI Can Ship

1. **UI** (Phase 5) — Menu entries, keyboard shortcuts, merge warning dialog, Merge Across
2. **Copy/paste merge recreation** (Phase 4 remainder) — Clipboard carries merge metadata, paste recreates merge at destination
3. **XLSX export** (Phase 6 remainder) — Write `<mergeCells>` element with per-cell border decomposition
4. **CSV export** (Phase 6 remainder) — Value in top-left only

### Test Coverage

34 merge-specific tests across the engine:

**`sheet.rs` (27 tests):** `test_merged_region_basic`, `test_merge_index_lookup`, `test_add_merge_overlap_rejected`, `test_add_merge_degenerate_ignored`, `test_remove_merge`, `test_merge_insert_row_shift`, `test_merge_insert_row_expand`, `test_merge_insert_row_below`, `test_merge_delete_row_shrink`, `test_merge_delete_row_degenerate`, `test_merge_delete_row_above`, `test_merge_insert_col_shift`, `test_merge_insert_col_expand`, `test_merge_delete_col_shrink`, `test_merge_delete_band_clips_top`, `test_merge_delete_band_clips_bottom`, `test_merge_delete_band_degenerates_wide`, `test_merge_delete_col_band_degenerates`, `test_merge_serde_roundtrip`, `test_resolve_merge_borders_no_borders`, `test_resolve_merge_borders_uniform_thin`, `test_resolve_merge_borders_mixed_precedence`, `test_resolve_merge_borders_single_styled_cell`, `test_adjacent_merges_shared_border`, `test_merge_at_viewport_boundary`, `test_merge_origin_coord`, `test_get_merge_returns_canonical_origin`

**`sheet.rs` semantic correctness (8 tests):** `test_set_value_redirects_to_origin`, `test_clear_cell_redirects_to_origin`, `test_set_value_preserves_hidden_style`, `test_set_format_no_redirect`, `test_clear_cell_preserves_hidden_style`, `test_set_cycle_error_preserves_hidden_style`, `test_range_clear_preserves_hidden_styles`

**`workbook.rs` formula tests (4 tests):** `test_formula_single_ref_redirect`, `test_formula_range_hidden_empty`, `test_formula_explicit_refs_redirect`, `test_formula_cross_sheet_ref_redirect`

**`cell.rs` merge override tests (3 tests):** `test_merge_override_empty`, `test_merge_override_replaces_fields`, `test_merge_override_clears_option_fields`

### Guarded Operations (12 total)

All bulk operations that could corrupt merge state are guarded via `block_if_merged()` (centralized helper in `app.rs`) or `paste_would_split_merge()` (specific partial-overlap detection in `clipboard.rs`):

| Operation | Guard | File |
|-----------|-------|------|
| Sort | `block_if_merged("sort")` | `sort_filter.rs` |
| AutoFilter | `block_if_merged("enable AutoFilter")` | `sort_filter.rs` |
| Fill Down | `block_if_merged("fill down")` | `fill.rs` |
| Fill Right | `block_if_merged("fill right")` | `fill.rs` |
| Fill Handle (vertical) | `block_if_merged("fill")` | `fill.rs` |
| Fill Handle (horizontal) | `block_if_merged("fill")` | `fill.rs` |
| Trim Whitespace | `block_if_merged("trim")` | `fill.rs` |
| Replace All | `block_if_merged("replace all")` | `find_replace.rs` |
| Fill Selection from Primary | `block_if_merged("fill selection")` | `editing.rs` |
| Paste | `paste_would_split_merge()` | `clipboard.rs` |
| Paste Values | `paste_would_split_merge()` | `clipboard.rs` |
| Paste Formulas | `paste_would_split_merge()` | `clipboard.rs` |

## Version History

| Version | Changes |
|---------|---------|
| v6 | Phase 2 complete (overlay rendering, spacers, spill exclusion, 17 tests). Phase 3 complete (merge-aware navigation: `find_data_boundary()` treats merges as data units, `jump_selection()`/`extend_jump_selection()` start from merge edges and snap/expand on landing, `confirm_goto()` redirects to origin, tab-chain Enter snaps to origin). 34 merge-specific tests. |
| v5 | Implementation: Phase 1 complete (data model, O(1) lookup, add/remove/normalize, insert/delete adjustment, serde). Phase 4 semantic correctness (value write redirect via `merge_origin_coord()`, formula CellRef redirect via `CellLookup` trait, 12 bulk op guards via `block_if_merged()`/`paste_would_split_merge()`). Phase 7 complete (sort/filter block, overlap detection, formula semantics). Phase 6 partial (.vgrid + XLSX import). 33 merge-specific tests. |
| v4 | Final review: debug_assert_no_merge_overlap invariant, dynamic range function note (OFFSET/INDIRECT follow same empty-cell rule), ResolvedMergeBorders cache for wide merges, single-cell paste clarified to work from any cell inside merge, freeze boundary flagged as Excel-divergent |
| v3 | Second review: border model corrected to per-side edge composition with precedence (not origin-only), formula error semantics for hidden cells, Merge Across concrete storage rules, precise insert/delete using grid-line indices, strict paste rules (block partial-intersection), text wrapping v1 limitation noted, freeze boundary hit-testing/selection defined, large-merge threshold (50k cells), hidden cell style storage strategy (keep style_id per cell), XLSX import overlap policy (drop + report) |
| v2 | First review: O(1) merge lookup via HashMap index, macOS Cmd+M conflict resolved (→ Ctrl+Shift+M), edge-sampling border rendering, set_value redirect to origin, unmerge inherits origin format, single-cell vs range formula semantics, normalize_merges() for insert/delete, simplified sort/filter (block with error), merge-aware layout (hit-testing, selection clamping, CenterAcross precedence), XLSX border composition on import/export |
| v1 | Spec created |
