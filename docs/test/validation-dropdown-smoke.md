# Validation Dropdown Smoke Test

Manual test script for verifying validation dropdown behavior. Run through after any changes to validation, dropdown, or event routing code.

## Setup Workbook

1. **Create list source values in Sheet2!A1:A5:**
   - A1: `Open`
   - A2: `In Progress`
   - A3: `Closed`
   - A4: `On Hold`
   - A5: `Blocked`

2. **Apply validation to Sheet1!B2:B20:**
   - List source: `=Sheet2!A1:A5`
   - Show dropdown: `true`
   - Ignore blank: `false`

   *(Until dialog exists: set via engine API or test fixture)*

---

## Keyboard Behavior

| # | Step | Expected Result |
|---|------|-----------------|
| 1 | Select B2 → Press Alt+Down | Dropdown opens, shows all 5 items |
| 2 | Type "pro" | Filters to "In Progress" only |
| 3 | Press Down/Up arrows | Highlight moves (even with 1 item, should not crash) |
| 4 | Press Enter | Value commits, dropdown closes, cell shows "In Progress" |
| 5 | Press Alt+Down again | Dropdown opens |
| 6 | Press Escape | Dropdown closes, value unchanged ("In Progress") |
| 7 | Press Alt+Down | Dropdown opens |
| 8 | Press Tab | Dropdown closes, selection moves to C2 (grid handles Tab) |

---

## Invalidation Behavior

| # | Step | Expected Result |
|---|------|-----------------|
| 1 | Open dropdown on B2 | Dropdown visible |
| 2 | Scroll grid with mouse wheel | Dropdown closes |
| 3 | Open dropdown again | Dropdown visible |
| 4 | Press Ctrl/Cmd + or - (zoom) | Dropdown closes |
| 5 | Open dropdown again | Dropdown visible |
| 6 | Press Ctrl/Cmd+F (Find) | Dropdown closes, Find dialog opens |
| 7 | Escape Find, open dropdown | Dropdown visible |
| 8 | Press Ctrl/Cmd+Shift+P (Command Palette) | Dropdown closes, palette opens |
| 9 | Escape palette, open dropdown | Dropdown visible |
| 10 | Press Left/Right arrow | Dropdown closes, selection moves |

---

## Fingerprint Invalidation (Critical)

| # | Step | Expected Result |
|---|------|-----------------|
| 1 | Open dropdown on B2 | Shows 5 items |
| 2 | Navigate to Sheet2!A2 | Leave dropdown open |
| 3 | Edit A2 from "In Progress" to "In Review" | |
| 4 | Return to B2 or trigger re-render | Dropdown must close (SourceChanged) or refresh |
| 5 | Open dropdown again | Should now show "In Review" instead of "In Progress" |

---

## Edge Cases

| # | Step | Expected Result |
|---|------|-----------------|
| 1 | Open dropdown, type "zzz" | Shows "0 items" in footer, no items visible |
| 2 | Press Enter with no matches | Nothing happens (no commit, dropdown stays open) |
| 3 | Press Escape | Dropdown closes |
| 4 | Open dropdown on cell with empty source list | Status bar shows "Validation list is empty" (or dropdown doesn't open) |
| 5 | Apply validation with >10,000 items | Footer shows "(truncated)", dropdown still works |

---

## Mouse Behavior

| # | Step | Expected Result |
|---|------|-----------------|
| 1 | Open dropdown, click on "Closed" | Value commits to "Closed", dropdown closes |
| 2 | Open dropdown | Dropdown visible |
| 3 | Click outside dropdown (on another cell) | Dropdown closes, selection does NOT change |
| 4 | Click outside dropdown (on grid background) | Dropdown closes |

---

## Pass Criteria

- [ ] All keyboard behaviors work as expected
- [ ] All invalidation triggers close the dropdown
- [ ] Fingerprint change closes/refreshes dropdown
- [ ] Edge cases handled gracefully
- [ ] Mouse click outside closes without side effects
- [ ] No console errors or panics

---

## Regression Notes

If any test fails, check:
1. Event routing order in `views/mod.rs` on_key_down
2. Action handler guards (MoveUp, MoveDown, PageUp, PageDown, ConfirmEdit, CancelEdit)
3. `close_validation_dropdown` calls in navigation.rs, dialogs.rs, sheet_ops.rs
4. Backdrop click handler in `validation_dropdown_view.rs` (must call `cx.stop_propagation()`)
