use gpui::*;
use crate::app::Spreadsheet;
use crate::actions::*;
use crate::mode::Mode;

pub(crate) fn bind(
    el: Div,
    cx: &mut Context<Spreadsheet>,
) -> Div {
    el
        // Navigation actions (formula mode: insert references, edit mode: move cursor, nav mode: move selection)
        .on_action(cx.listener(|this, _: &MoveUp, _, cx| {
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AskAI) {
                return;
            }
            // Validation dropdown navigation takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.move_up();
                    cx.notify();
                }
                return;
            }
            // Lua console: history prev
            if this.lua_console.visible {
                this.lua_console.history_prev();
                cx.notify();
                return;
            }
            // Formula mode: Point submode does ref-pick, Caret submode is no-op for Up/Down
            if this.mode.is_formula() {
                if this.formula_is_caret_mode() {
                    // Caret mode: Up/Down are no-op (single-line editor)
                    return;
                }
                this.formula_move_ref(-1, 0, cx);
                return;
            }
            // Autocomplete navigation (Edit mode only, not Formula)
            if this.autocomplete_visible {
                this.autocomplete_up(cx);
                return;
            }
            // Arrow keys break the tab-chain (only Tab/Enter preserve it)
            this.nav_perf.mark_key_action();
            this.tab_chain_origin_col = None;
            match this.mode {
                Mode::Command => this.palette_up(cx),
                Mode::FontPicker => this.font_picker_up(cx),
                Mode::ThemePicker => this.theme_picker_up(cx),
                // Edit mode: commit-on-arrow (fast data entry, Excel-like)
                Mode::Edit => this.confirm_edit_up(cx),
                _ => {
                    // Batch: accumulate for flush at render start
                    this.pending_nav_dy -= 1;
                    cx.notify();
                }
            }
        }))
        .on_action(cx.listener(|this, _: &MoveDown, _, cx| {
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AskAI) {
                return;
            }
            // Validation dropdown navigation takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.move_down();
                    cx.notify();
                }
                return;
            }
            // Lua console: history next
            if this.lua_console.visible {
                this.lua_console.history_next();
                cx.notify();
                return;
            }
            // Formula mode: Point submode does ref-pick, Caret submode is no-op for Up/Down
            if this.mode.is_formula() {
                if this.formula_is_caret_mode() {
                    // Caret mode: Up/Down are no-op (single-line editor)
                    return;
                }
                this.formula_move_ref(1, 0, cx);
                return;
            }
            // Autocomplete navigation (Edit mode only, not Formula)
            if this.autocomplete_visible {
                this.autocomplete_down(cx);
                return;
            }
            // Arrow keys break the tab-chain (only Tab/Enter preserve it)
            this.nav_perf.mark_key_action();
            this.tab_chain_origin_col = None;
            match this.mode {
                Mode::Command => this.palette_down(cx),
                Mode::FontPicker => this.font_picker_down(cx),
                Mode::ThemePicker => this.theme_picker_down(cx),
                // Edit mode: commit-on-arrow (fast data entry, Excel-like)
                Mode::Edit => this.confirm_edit(cx),
                _ => {
                    // Batch: accumulate for flush at render start
                    this.pending_nav_dy += 1;
                    cx.notify();
                }
            }
        }))
        .on_action(cx.listener(|this, _: &MoveLeft, window, cx| {
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AskAI) {
                return;
            }
            // Lua console: cursor left
            if this.lua_console.visible {
                this.lua_console.cursor_left();
                cx.notify();
                return;
            }
            // Modal modes: capture arrow keys, don't leak to grid
            // List-only modals (no text input): arrows are no-op for left/right
            if matches!(this.mode, Mode::ThemePicker | Mode::FontPicker) {
                return;  // No horizontal navigation in vertical lists
            }
            // Text input modals: move cursor left in the input field
            if matches!(this.mode, Mode::Command | Mode::GoTo | Mode::Find |
                Mode::RenameSymbol | Mode::CreateNamedRange | Mode::EditDescription |
                Mode::ExtractNamedRange | Mode::License) {
                // These modes handle cursor movement in their on_key_down handlers
                // Just block the event from reaching the grid
                return;
            }
            if this.mode.is_formula() {
                if this.formula_is_caret_mode() {
                    // Caret mode: move cursor in formula text
                    this.move_edit_cursor_left(cx);
                    this.update_edit_scroll(window);
                } else {
                    // Point mode: ref-picking
                    this.formula_move_ref(0, -1, cx);
                }
            } else if this.mode.is_editing() {
                // Edit mode: commit-on-arrow (fast data entry, Excel-like)
                this.nav_perf.mark_key_action();
                this.tab_chain_origin_col = None;  // Arrow breaks tab chain
                this.confirm_edit_and_move_left(cx);
            } else {
                this.nav_perf.mark_key_action();
                this.tab_chain_origin_col = None;  // Arrow breaks tab chain
                // Batch: accumulate for flush at render start
                this.pending_nav_dx -= 1;
                cx.notify();
            }
        }))
        .on_action(cx.listener(|this, _: &MoveRight, window, cx| {
            // Let AI dialogs handle their own keys
            if matches!(this.mode, Mode::AISettings | Mode::AskAI) {
                return;
            }
            // Lua console: cursor right
            if this.lua_console.visible {
                this.lua_console.cursor_right();
                cx.notify();
                return;
            }
            // Modal modes: capture arrow keys, don't leak to grid
            // List-only modals (no text input): arrows are no-op for left/right
            if matches!(this.mode, Mode::ThemePicker | Mode::FontPicker) {
                return;  // No horizontal navigation in vertical lists
            }
            // Text input modals: move cursor right in the input field
            if matches!(this.mode, Mode::Command | Mode::GoTo | Mode::Find |
                Mode::RenameSymbol | Mode::CreateNamedRange | Mode::EditDescription |
                Mode::ExtractNamedRange | Mode::License) {
                // These modes handle cursor movement in their on_key_down handlers
                // Just block the event from reaching the grid
                return;
            }
            if this.mode.is_formula() {
                if this.formula_is_caret_mode() {
                    // Caret mode: move cursor in formula text
                    this.move_edit_cursor_right(cx);
                    this.update_edit_scroll(window);
                } else {
                    // Point mode: ref-picking
                    this.formula_move_ref(0, 1, cx);
                }
            } else if this.mode.is_editing() {
                // Edit mode: commit-on-arrow (fast data entry, Excel-like)
                this.nav_perf.mark_key_action();
                this.tab_chain_origin_col = None;  // Arrow breaks tab chain
                this.confirm_edit_and_move_right(cx);
            } else {
                this.nav_perf.mark_key_action();
                this.tab_chain_origin_col = None;  // Arrow breaks tab chain
                // Batch: accumulate for flush at render start
                this.pending_nav_dx += 1;
                cx.notify();
            }
        }))
        .on_action(cx.listener(|this, _: &JumpUp, _, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(-1, 0, cx);
            } else {
                this.jump_selection(-1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &JumpDown, _, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(1, 0, cx);
            } else {
                this.jump_selection(1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &JumpLeft, window, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.move_edit_cursor_word_left(cx);
            } else {
                this.jump_selection(0, -1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &JumpRight, window, cx| {
            if this.mode.is_formula() {
                this.formula_jump_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.move_edit_cursor_word_right(cx);
            } else {
                this.jump_selection(0, 1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &MoveToStart, _, cx| {
            this.view_state.selected = (0, 0);
            this.view_state.scroll_row = 0;
            this.view_state.scroll_col = 0;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &MoveToEnd, _, cx| {
            this.view_state.selected = (crate::app::NUM_ROWS - 1, crate::app::NUM_COLS - 1);
            this.view_state.scroll_row = crate::app::NUM_ROWS.saturating_sub(this.visible_rows());
            this.view_state.scroll_col = crate::app::NUM_COLS.saturating_sub(this.visible_cols());
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &PageUp, _, cx| {
            // Validation dropdown takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.page_up(10);
                    cx.notify();
                }
                return;
            }
            this.page_up(cx);
        }))
        .on_action(cx.listener(|this, _: &PageDown, _, cx| {
            // Validation dropdown takes priority
            if this.is_validation_dropdown_open() {
                if let Some(state) = this.validation_dropdown.as_open_mut() {
                    state.page_down(10);
                    cx.notify();
                }
                return;
            }
            this.page_down(cx);
        }))
        // Selection extension (formula mode: extend range reference)
        .on_action(cx.listener(|this, _: &ExtendUp, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(-1, 0, cx);
            } else {
                this.extend_selection(-1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendDown, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(1, 0, cx);
            } else {
                this.extend_selection(1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendLeft, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_left(cx);
            } else {
                this.extend_selection(0, -1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &ExtendRight, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_right(cx);
            } else {
                this.extend_selection(0, 1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpUp, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(-1, 0, cx);
            } else {
                this.extend_jump_selection(-1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpDown, _, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(1, 0, cx);
            } else {
                this.extend_jump_selection(1, 0, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpLeft, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(0, -1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_word_left(cx);
            } else {
                this.extend_jump_selection(0, -1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &ExtendJumpRight, window, cx| {
            if this.mode.is_formula() {
                this.formula_extend_jump_ref(0, 1, cx);
            } else if this.mode == Mode::Edit {
                this.select_edit_cursor_word_right(cx);
            } else {
                this.extend_jump_selection(0, 1, cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &SelectAll, window, cx| {
            if this.mode == Mode::ColorPicker {
                this.color_picker_select_all(cx);
                return;
            }
            if this.mode == Mode::Edit {
                this.select_all_edit(cx);
            } else {
                this.select_all(cx);
            }
            this.update_edit_scroll(window);
        }))
        .on_action(cx.listener(|this, _: &SelectBlanks, _, cx| {
            if !this.mode.is_editing() {
                this.select_blanks(cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectRow, _, cx| {
            if !this.mode.is_editing() {
                let row = this.view_state.selected.0;
                this.select_row(row, false, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &SelectColumn, _, cx| {
            if !this.mode.is_editing() {
                let col = this.view_state.selected.1;
                this.select_col(col, false, cx);
            }
        }))
        // Go To dialog
        .on_action(cx.listener(|this, _: &GoToCell, _, cx| {
            this.show_goto(cx);
        }))
        // Jump to active cell (Ctrl+Backspace)
        .on_action(cx.listener(|this, _: &JumpToActiveCell, _, cx| {
            this.ensure_visible(cx);
            cx.notify();
        }))
        // Find dialog
        .on_action(cx.listener(|this, _: &FindInCells, _, cx| {
            this.show_find(cx);
        }))
        .on_action(cx.listener(|this, _: &FindNext, _, cx| {
            this.find_next(cx);
        }))
        .on_action(cx.listener(|this, _: &FindPrev, _, cx| {
            this.find_prev(cx);
        }))
        .on_action(cx.listener(|this, _: &FindReplace, _, cx| {
            this.show_find_replace(cx);
        }))
        .on_action(cx.listener(|this, _: &ReplaceNext, _, cx| {
            this.replace_next(cx);
        }))
        .on_action(cx.listener(|this, _: &ReplaceAll, _, cx| {
            this.replace_all(cx);
        }))
        // IDE-style navigation (Find References / Go to Precedents)
        .on_action(cx.listener(|this, _: &FindReferences, _, cx| {
            // In edit mode, check if cursor is on a named range
            if this.mode.is_editing() {
                if let Some(name) = this.named_range_at_cursor(cx) {
                    this.show_named_range_references(&name, cx);
                    return;
                }
            }
            // Fall back to cell references
            let (row, col) = this.view_state.selected;
            this.show_references(row, col, cx);
        }))
        .on_action(cx.listener(|this, _: &GoToPrecedents, _, cx| {
            // In edit mode, check if cursor is on a named range
            if this.mode.is_editing() {
                if let Some(name) = this.named_range_at_cursor(cx) {
                    this.go_to_named_range_definition(&name, cx);
                    return;
                }
            }
            // Fall back to cell precedents
            let (row, col) = this.view_state.selected;
            this.show_precedents(row, col, cx);
        }))
        .on_action(cx.listener(|this, _: &RenameSymbol, _, cx| {
            this.show_rename_symbol(None, cx);
        }))
        .on_action(cx.listener(|this, _: &CreateNamedRange, _, cx| {
            this.show_create_named_range(cx);
        }))
        // Validation failure navigation (F8 / Shift+F8)
        .on_action(cx.listener(|this, _: &NextInvalidCell, _, cx| {
            this.next_invalid_cell(cx);
        }))
        .on_action(cx.listener(|this, _: &PrevInvalidCell, _, cx| {
            this.prev_invalid_cell(cx);
        }))
}
