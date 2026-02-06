# Border Feature QA Checklist

Manual sanity pass for border formatting changes. Run through before any formatting release.

## Color Selection

- [ ] Click swatch **Red** → Click **Outline** → Red outline appears on selection
- [ ] Click **Auto** → Click **Outline** → Uses theme default line color (dark gray)
- [ ] Click **...** → Enter hex `#FF8000` → Click **Apply** → Orange borders appear
- [ ] Color swatches show selected state (larger size, ring) when active

## Border Application

- [ ] **All Borders** → Every cell edge has borders
- [ ] **Outline Border** → Only perimeter edges, interior untouched
- [ ] **Inside Borders** → Only internal grid lines, no outline
- [ ] **Clear Borders** → All borders removed from selection
- [ ] **Top/Bottom/Left/Right** → Single edge applied correctly

## Inside Borders Edge Cases

- [ ] Select 1×5 range → **Inside** button is disabled (no internal edges)
- [ ] Select 5×1 range → **Inside** button is disabled
- [ ] Select 2×2 range → **Inside** creates cross pattern (1 horizontal + 1 vertical internal line)
- [ ] Select 3×3 range → **Inside** creates full internal grid (outline untouched)

## Persistence

- [ ] Apply colored borders → Save → Close → Reopen → Colors match exactly
- [ ] Auto (None) borders → Save → Load → Still shows theme default (not literal color)
- [ ] Different colors on different edges → All persist after round-trip

## Copy/Paste

- [ ] Copy bordered range → Paste → Borders included (with colors)
- [ ] Cut bordered range → Paste → Borders move to new location

## Color Picker Integration

- [ ] **...** button opens color picker with "Border Color" title
- [ ] Selecting color in picker updates `current_border_color`
- [ ] Esc closes picker, returns focus to inspector panel
- [ ] Recent colors section shows previously used border colors

## Tooltips

- [ ] Hover **Auto** → "Automatic (theme default)"
- [ ] Hover **...** → "More colors..."
- [ ] Hover color swatch → Shows color name (Red, Blue, etc.)
- [ ] Hover border icon → Shows action (Outline Border, All Borders, etc.)

---

**Design decision documented:**

> v1 simplification: When rendering borders, the first non-None color found
> (checking top, right, bottom, left in order) is used for all edges of that cell.
> Per-edge colors are stored and persisted, but rendered uniformly.
> Per-edge color rendering may come in a future version.

See `gpui-app/src/views/grid.rs` for the rendering logic.
