# Formula Bar Editing

Make the formula bar a real editor surface with mouse interaction and proper popup placement.

## Recent Changes

**Phase 2 Complete** (drag selection + auto-scroll):
- Added `formula_bar_drag_anchor: Option<usize>` for tracking drag state
- Mouse handlers: `on_mouse_down`, `on_mouse_move`, `on_mouse_up`
- Selection rendered as background rect overlay (respects scroll offset)
- Auto-scroll when dragging near edges (10px margin, 4px/frame speed)
- Centralized layout constants: `FORMULA_BAR_CELL_REF_WIDTH`, `FORMULA_BAR_FX_WIDTH`, `FORMULA_BAR_PADDING`, `FORMULA_BAR_TEXT_LEFT`

**Phase 1 Complete** (click-to-place caret):
- `formula_bar_text_rect` calculated during render (window coordinates)
- Fixed coordinate space mismatch (was causing caret to always land at end)
- Boundary cache for hit-testing with `partition_point`

## Current State

- Formula bar displays text and accepts keyboard input
- Caret positioning works (fixed in recent commit)
- **Phase 1 DONE**: Click-to-place caret with scroll support
- **Phase 2 DONE**: Drag-to-select with auto-scroll near edges
- **TODO**: Proper popup anchoring (Phase 3)

---

## Implementation Status

| Phase | Status | Description |
|-------|--------|-------------|
| Phase 1 | ✅ Done | Click caret placement + scroll |
| Phase 2 | ✅ Done | Drag selection + auto-scroll |
| Phase 3 | ⏳ Next | Fix popup blocking |
| Phase 4 | 📋 Planned | Enhanced selection (double-click, shift-click) |

---

## Architecture Decisions

### 1. Drag State: `Option<usize>` not separate bool

```rust
// In Spreadsheet:
formula_bar_drag_anchor: Option<usize>  // None = not dragging, Some(byte) = anchor
```

- `None` = not dragging
- `Some(anchor)` = dragging; cursor updates extend selection from anchor
- Reduces state drift, single thing to clear

### 2. Cache Boundary Positions (not ShapedLine) ✅ Implemented

`ShapedLine` may carry references or not be Send/Sync. Cache only what's needed:

```rust
// In Spreadsheet:
formula_bar_char_boundaries: Vec<usize>  // byte offsets: [0, 1, 2, ..., len]
formula_bar_boundary_xs: Vec<f32>        // x positions aligned to boundaries
formula_bar_text_width: f32              // last boundary x (for scroll calc)
formula_bar_cache_dirty: bool            // dirty flag for lazy rebuild
```

**On buffer change** (in `start_edit`, `insert_char`, `backspace`, etc.):
- Set `formula_bar_cache_dirty = true`
- Rebuild lazily via `maybe_rebuild_formula_bar_cache(window)`

### 3. Click Hit-Testing: Find Closest Boundary ✅ Implemented

Uses `partition_point` + closest-boundary comparison for "sticky correct" feel:

```rust
fn byte_index_for_x(&self, x: f32) -> usize {
    let boundaries = &self.formula_bar_char_boundaries;
    let xs = &self.formula_bar_boundary_xs;

    if boundaries.is_empty() || xs.is_empty() {
        return 0;
    }

    // Find first boundary whose x >= click_x
    let i = xs.partition_point(|&bx| bx < x);

    let right_idx = i.min(boundaries.len() - 1);
    let left_idx = i.saturating_sub(1);

    let right = boundaries[right_idx];
    let left = boundaries[left_idx];

    let xr = xs[right_idx];
    let xl = xs[left_idx];

    // Return whichever boundary is closer
    if (x - xl).abs() <= (xr - x).abs() { left } else { right }
}
```

### 4. Editor Surface Enum ✅ Implemented

```rust
#[derive(Clone, Copy, PartialEq, Default)]
pub enum EditorSurface {
    #[default]
    Cell,
    FormulaBar,
}

// In Spreadsheet:
pub active_editor: EditorSurface
```

**Transitions**:
| Event | New State |
|-------|-----------|
| Click in cell while editing | `Cell` |
| Click in formula bar | `FormulaBar` |
| F2 to start edit | `Cell` |
| Esc / cancel edit | `Cell` |
| Enter / confirm edit | `Cell` |
| Click outside while editing | `Cell` (ends edit) |

### 5. Selection Rendering: Overlay Rect with Scroll/Origin

Don't bake selection into TextRuns. Draw rect behind text.

**Coordinate math** (must include scroll + origin):
```rust
// In formula bar render:
let origin_x = text_area_left;  // after "fx" button + padding
let scroll_x = self.formula_bar_scroll_x;

// Selection rect screen coordinates
let sel_start_x = origin_x + scroll_x + boundary_xs[sel_start_boundary];
let sel_end_x = origin_x + scroll_x + boundary_xs[sel_end_boundary];

// Clip to formula bar text rect
let sel_left = sel_start_x.max(text_area_left);
let sel_right = sel_end_x.min(text_area_right);
```

**Render order**:
1. Selection background rect (highlight color)
2. Styled text on top
3. Caret overlay

### 6. Popup Anchoring: Hard Constraint

**Rule**: When `active_editor == FormulaBar`, popup MUST NOT overlap formula bar. Period.

```rust
fn popup_position(&self, popup_height: f32) -> Point {
    let gap = 6.0;

    let (anchor_rect, forbidden_rect) = match self.active_editor {
        EditorSurface::FormulaBar => {
            (self.formula_bar_rect, self.formula_bar_rect)
        }
        EditorSurface::Cell => {
            (self.active_cell_rect, self.active_cell_rect)
        }
    };

    // Preferred: below anchor
    let mut y = anchor_rect.bottom() + gap;

    // Flip above if no room below
    if y + popup_height > viewport.bottom() {
        y = anchor_rect.top() - gap - popup_height;
    }

    // Hard constraint: never overlap forbidden rect
    if y < forbidden_rect.bottom() && y + popup_height > forbidden_rect.top() {
        y = forbidden_rect.bottom() + gap;
    }

    Point { x: anchor_rect.left(), y }
}
```

No pointer-events tricks needed—just don't overlap.

### 7. Auto-Scroll While Dragging ✅ Implemented

```rust
// In on_mouse_move handler:
if this.formula_bar_drag_anchor.is_some() {
    let edge_margin = 10.0;
    let scroll_speed = 4.0;

    if mouse_x < text_left + edge_margin {
        // Near left edge - scroll to show content on the left
        this.formula_bar_scroll_x = (this.formula_bar_scroll_x + scroll_speed).min(0.0);
    } else if mouse_x > text_right - edge_margin {
        // Near right edge - scroll to show content on the right
        let max_scroll = -(this.formula_bar_text_width).max(0.0);
        this.formula_bar_scroll_x = (this.formula_bar_scroll_x - scroll_speed).max(max_scroll);
    }
}
```

---

## Gotchas (Don't Step on These Rakes)

### 1. Cache Rebuild: Use Dirty Flag ✅ Implemented

Don't rebuild on every keystroke—shaping is expensive on long formulas.

```rust
formula_bar_cache_dirty: bool

// Set dirty on buffer change (in insert_char, backspace, etc.)
self.formula_bar_cache_dirty = true;

// Rebuild once, centralized (before mouse handling)
this.maybe_rebuild_formula_bar_cache(window);
```

Same pattern as `edit_scroll_dirty`.

### 2. Font Settings Must Match Render ✅ Implemented

Cache uses `px(14.0)` and default font. Formula bar render uses `.text_sm()` which is 14px.

**Centralized layout constants** (in `app.rs`, single source of truth):
```rust
pub const FORMULA_BAR_CELL_REF_WIDTH: f32 = 60.0;
pub const FORMULA_BAR_FX_WIDTH: f32 = 30.0;
pub const FORMULA_BAR_PADDING: f32 = 8.0;  // px_2
pub const FORMULA_BAR_TEXT_LEFT: f32 = 98.0;  // computed from above
```

Used in both render code (`formula_bar.rs`) and rect calculation (`app.rs`).

### 3. Debug Assert for Monotonic xs ✅ Implemented

`partition_point` assumes monotonic. Ligatures could break this (rare).

```rust
debug_assert!(
    boundary_xs.windows(2).all(|w| w[0] <= w[1] + 0.01),
    "boundary_xs not monotonic - text shaping issue"
);
```

### 4. Popup x: Align to Caret (Polish)

For formula bar, anchor popup to caret x, not left edge. Optional for v1.

### 5. Click Outside: Don't End Edit on Popup Click

"Click outside while editing ends edit" — but clicking autocomplete/signature help shouldn't cancel.

**Rule**:
- Clicks on autocomplete/signature help → don't end edit
- Clicks outside editor surfaces AND outside popups → end edit

### 6. Auto-Scroll: Cap Speed ✅ Implemented

Using 4px per mousemove with 10px edge margin. Clamped to valid scroll range.

---

## Implementation Phases

### Phase 1: Click Caret + Scroll ✅ DONE

**Implemented in commit `37685c3` + follow-up**

1. ✅ Added state:
   ```rust
   active_editor: EditorSurface
   formula_bar_scroll_x: f32
   formula_bar_text_rect: Bounds<Pixels>  // Text area in window coords (for hit-testing)
   formula_bar_cache_dirty: bool
   formula_bar_char_boundaries: Vec<usize>
   formula_bar_boundary_xs: Vec<f32>
   formula_bar_text_width: f32
   ```

2. ✅ Added `rebuild_formula_bar_cache()` + `maybe_rebuild_formula_bar_cache()`

3. ✅ Added `byte_index_for_x()` using partition_point + closest boundary

4. ✅ Added mouse handler in `formula_bar.rs`:
   - Starts edit if not editing
   - Rebuilds cache if dirty
   - Converts click x to byte index
   - Sets cursor and clears selection
   - Sets `active_editor = FormulaBar`
   - Ensures caret visible

5. ✅ Implemented `ensure_formula_bar_caret_visible()`

6. ✅ Applied scroll offset to formula bar text and caret rendering

7. ✅ Fixed coordinate space mismatch:
   - `formula_bar_text_rect` calculated during render (in window coords)
   - Mouse handler uses `event.position.x - text_rect.origin.x - scroll_x`
   - Both are in window coordinates now (no more local vs global mismatch)

### Phase 2: Drag Selection ✅ DONE

1. ✅ Added drag tracking state:
   ```rust
   formula_bar_drag_anchor: Option<usize>
   ```

2. ✅ Mouse handlers:
   - `on_mouse_down`: sets `formula_bar_drag_anchor` + `edit_selection_anchor`
   - `on_mouse_move`: extends selection if dragging, with auto-scroll near edges
   - `on_mouse_up`: clears drag anchor, clears selection if click (no drag)

3. ✅ Selection rendered as background rect overlay (before text, after scroll offset)

### Phase 3: Fix Popup Blocking

1. Use `active_editor` to determine anchor
2. Hard constraint: popup below forbidden_rect.bottom() when overlap detected
3. Reset `active_editor` to `Cell` on Esc/Enter/click-outside

### Phase 4: Enhanced Selection (Nice to Have)

- Double-click: select word
- Triple-click: select all
- Shift+Click: extend selection
- Ctrl/Cmd+A: select all (in formula bar context)
- Alt+Arrow: word jump
- ~~Auto-scroll while dragging~~ (done in Phase 2)

---

## File Changes

| File | Changes | Status |
|------|---------|--------|
| `app.rs` | `EditorSurface`, centralized layout constants, scroll/cache state, `formula_bar_text_rect`, `formula_bar_drag_anchor`, `byte_index_for_x()`, `rebuild_formula_bar_cache()` | ✅ Done |
| `views/formula_bar.rs` | Mouse handlers (down/move/up), selection rendering, auto-scroll, scroll offset rendering | ✅ Done |
| `views/mod.rs` | Popup placement based on `active_editor` | ⏳ Phase 3 |

---

## Testing Checklist

### Phase 1 ✅
- [x] Click places caret at correct position (sticky to nearest char)
- [ ] Click at far right of long formula → scrolls to reveal caret
- [x] Click at position 0 → caret at start
- [x] Click between two chars → lands on closest one

### Phase 2 ✅
- [x] Drag selects correct range
- [x] Selection renders with highlight rect (not colored text runs)
- [x] Drag beyond visible region → auto-scroll
- [ ] Typing replaces selection (uses existing `edit_selection_anchor` logic)
- [ ] Backspace/Delete removes selection (uses existing logic)
- [ ] Copy/Cut/Paste work with selection (uses existing logic)

### Phase 3
- [ ] Popup appears below formula bar when `active_editor == FormulaBar`
- [ ] Popup visible + click formula bar → formula bar receives click
- [ ] Popup never overlaps formula bar input area
- [ ] Esc resets `active_editor` to `Cell`
- [ ] Enter resets `active_editor` to `Cell`

### Regression
- [ ] Cell editor still works correctly
- [ ] Keyboard navigation unchanged
- [ ] Formula mode ref selection still works
- [ ] Popup anchors to cell when editing in cell
