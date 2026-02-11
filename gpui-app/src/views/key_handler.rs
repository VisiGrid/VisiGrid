use gpui::*;
use crate::app::{Spreadsheet, CreateNameFocus};
use crate::mode::{Mode, InspectorTab};

pub(crate) fn handle_key_down(
    this: &mut Spreadsheet,
    event: &KeyDownEvent,
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) {
    // Format bar owns focus: don't route keys to the grid
    if this.ui.format_bar.is_active(window) {
        return;
    }

    // KeyTips: if active, route key to handler (takes precedence)
    if this.keytips_active {
        if this.keytips_handle_key(&event.keystroke.key, cx) {
            return;
        }
    }

    // Menu dropdown keyboard navigation
    if this.open_menu.is_some() {
        let key = event.keystroke.key.as_str();
        match key {
            "escape" => {
                this.close_menu(cx);
                return;
            }
            "up" => {
                this.menu_highlight_prev(cx);
                return;
            }
            "down" => {
                this.menu_highlight_next(cx);
                return;
            }
            "left" => {
                this.menu_switch_prev(cx);
                return;
            }
            "right" => {
                this.menu_switch_next(cx);
                return;
            }
            "enter" => {
                this.menu_execute_highlighted(window, cx);
                return;
            }
            _ => {
                // Non-modifier keys: try letter accelerator first, then close menu
                if key != "shift" && key != "control" && key != "alt" && key != "cmd" {
                    if key.len() == 1 {
                        let ch = key.chars().next().unwrap().to_ascii_lowercase();
                        if this.menu_execute_by_letter(ch, window, cx) {
                            return;
                        }
                    }
                    this.close_menu(cx);
                }
            }
        }
    }

    // Let AI dialogs handle their own keys
    if matches!(this.mode, Mode::AISettings | Mode::AiDialog) {
        return;
    }

    // Let Number Format Editor dialog handle its own keys
    if this.mode == Mode::NumberFormatEditor {
        return;
    }

    // Route keys to Lua console when visible
    if this.lua_console.visible {
        crate::views::lua_console::handle_console_key_from_main(this, event, window, cx);
        return;
    }

    // Close-confirm dialog traps all keyboard input
    if this.close_confirm_visible {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.close_confirm_visible = false;
                cx.notify();
            }
            "tab" => {
                if event.keystroke.modifiers.shift {
                    this.close_confirm_focused = if this.close_confirm_focused == 0 { 2 } else { this.close_confirm_focused - 1 };
                } else {
                    this.close_confirm_focused = (this.close_confirm_focused + 1) % 3;
                }
                cx.notify();
            }
            "enter" | "space" => {
                match this.close_confirm_focused {
                    0 => {
                        this.close_confirm_visible = false;
                        cx.notify();
                    }
                    1 => {
                        this.close_confirm_visible = false;
                        this.prepare_close(cx);
                        window.remove_window();
                    }
                    _ => {
                        this.close_confirm_visible = false;
                        let saved = this.save_and_close(cx);
                        if saved {
                            this.prepare_close(cx);
                            window.remove_window();
                        }
                    }
                }
            }
            _ => {}
        }
        return;
    }

    // Handle sheet context menu (close on any key)
    if this.sheet_context_menu.is_some() {
        this.hide_sheet_context_menu(cx);
        if event.keystroke.key == "escape" {
            return;
        }
    }

    // Handle cell/header context menu (close on non-modifier keys)
    if this.context_menu.is_some() {
        let key = &event.keystroke.key;
        if key != "shift" && key != "control" && key != "alt" && key != "cmd" {
            this.hide_context_menu(cx);
            if key == "escape" {
                return;
            }
        }
    }

    // Handle filter dropdown keys
    if this.filter_dropdown_col.is_some() {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.close_filter_dropdown(cx);
                return;
            }
            "enter" => {
                this.apply_filter_dropdown(cx);
                return;
            }
            "backspace" => {
                this.filter_search_text.pop();
                cx.notify();
                return;
            }
            _ => {
                // Type into search box
                if let Some(ch) = &event.keystroke.key_char {
                    if !event.keystroke.modifiers.control && !event.keystroke.modifiers.alt {
                        this.filter_search_text.push_str(ch);
                        cx.notify();
                        return;
                    }
                }
            }
        }
    }

    // Handle validation dropdown keys (MUST come before edit-mode triggers)
    if this.is_validation_dropdown_open() {
        let modifiers = crate::validation_dropdown::KeyModifiers {
            control: event.keystroke.modifiers.control,
            alt: event.keystroke.modifiers.alt,
            shift: event.keystroke.modifiers.shift,
            platform: event.keystroke.modifiers.platform,
        };

        // Route key through dropdown first
        if this.route_dropdown_key_event(&event.keystroke.key, modifiers, cx) {
            return;
        }

        // Handle character input for filtering
        if let Some(key_char) = &event.keystroke.key_char {
            for ch in key_char.chars() {
                if this.route_dropdown_char_event(ch, cx) {
                    return;
                }
            }
        }
    }

    // Cancel fill handle drag with Escape
    if this.is_fill_dragging() && event.keystroke.key == "escape" {
        this.cancel_fill_drag(cx);
        return;
    }

    // Exit zen mode with Escape
    if this.zen_mode && event.keystroke.key == "escape" {
        this.zen_mode = false;
        cx.notify();
        return;
    }

    // F1 hold-to-peek: show context help while F1 is held
    if event.keystroke.key == "f1" {
        this.f1_help_visible = true;
        cx.notify();
        return;
    }

    // Space hold-to-peek: preview workbook state before selected history entry
    if event.keystroke.key == "space"
        && !this.mode.is_editing()
        && !this.is_previewing()
        && this.selected_history_id.is_some()
        && this.history_highlight_range.is_some()
    {
        if let Err(e) = this.enter_preview(cx) {
            this.status_message = Some(format!("Preview failed: {}", e));
            cx.notify();
        }
        return;
    }

    // Space+Arrow scrubbing: while previewing, Up/Down navigates history entries
    // Makes the feature visceral: "scrub the timeline"
    if this.is_previewing() && (event.keystroke.key == "up" || event.keystroke.key == "down") {
        let direction = if event.keystroke.key == "up" { -1i32 } else { 1 };
        this.scrub_preview(direction, cx);
        return;
    }

    // Handle sheet rename mode
    if this.renaming_sheet.is_some() {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.cancel_sheet_rename(cx);
                return;
            }
            "enter" => {
                this.confirm_sheet_rename(cx);
                return;
            }
            "backspace" => {
                this.sheet_rename_backspace(cx);
                return;
            }
            "delete" => {
                this.sheet_rename_delete(cx);
                return;
            }
            "left" => {
                this.sheet_rename_cursor_left(cx);
                return;
            }
            "right" => {
                this.sheet_rename_cursor_right(cx);
                return;
            }
            "home" => {
                this.sheet_rename_cursor_home(cx);
                return;
            }
            "end" => {
                this.sheet_rename_cursor_end(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for rename
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars().filter(|c| !c.is_control()) {
                    this.sheet_rename_input_char(c, cx);
                }
                return;
            }
        }
        return;
    }

    // Handle Formula Autocomplete (highest priority when visible)
    if this.autocomplete_visible {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.autocomplete_dismiss(cx);
                return;
            }
            "enter" | "tab" => {
                this.autocomplete_accept(cx);
                return;
            }
            "up" => {
                this.autocomplete_up(cx);
                return;
            }
            "down" => {
                this.autocomplete_down(cx);
                return;
            }
            "shift-tab" => {
                // Dismiss autocomplete on Shift+Tab (spec: no accept)
                this.autocomplete_dismiss(cx);
                return;
            }
            _ => {
                // Other keys: let them pass through to normal handling
                // but the input will update autocomplete
            }
        }
    }

    // Handle Command Palette mode
    if this.mode == Mode::Command {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_palette(cx);
                return;
            }
            "enter" => {
                if event.keystroke.modifiers.control {
                    // Ctrl+Enter = secondary action (copy path, copy ref, show help)
                    this.palette_execute_secondary(window, cx);
                } else if event.keystroke.modifiers.shift {
                    // Shift+Enter = preview (apply without closing)
                    this.palette_preview(cx);
                } else {
                    // Plain Enter = execute and close
                    this.palette_execute(window, cx);
                }
                return;
            }
            "up" => {
                this.palette_up(cx);
                return;
            }
            "down" => {
                this.palette_down(cx);
                return;
            }
            "backspace" => {
                this.palette_backspace(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for palette
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars() {
                    this.palette_insert_char(c, cx);
                }
                return;
            }
        }
    }

    // Handle Font Picker mode
    if this.mode == Mode::FontPicker {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_font_picker(cx);
                return;
            }
            "enter" => {
                this.font_picker_execute(cx);
                return;
            }
            "up" => {
                this.font_picker_up(cx);
                return;
            }
            "down" => {
                this.font_picker_down(cx);
                return;
            }
            "backspace" => {
                this.font_picker_backspace(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for font picker
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars() {
                    this.font_picker_insert_char(c, cx);
                }
                return;
            }
        }
    }

    // Handle Color Picker mode
    if this.mode == Mode::ColorPicker {
        let has_modifier = event.keystroke.modifiers.control
            || event.keystroke.modifiers.alt
            || event.keystroke.modifiers.platform;
        let key_char = event.keystroke.key_char.as_deref();
        if this.color_picker_handle_key(
            event.keystroke.key.as_str(),
            key_char,
            has_modifier,
            window,
            cx,
        ) {
            return;
        }
    }

    // Handle Theme Picker mode
    if this.mode == Mode::ThemePicker {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_theme_picker(cx);
                return;
            }
            "enter" => {
                this.theme_picker_execute(window, cx);
                return;
            }
            "up" => {
                this.theme_picker_up(cx);
                return;
            }
            "down" => {
                this.theme_picker_down(cx);
                return;
            }
            "backspace" => {
                this.theme_picker_backspace(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for theme picker
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars() {
                    this.theme_picker_insert_char(c, cx);
                }
                return;
            }
        }
    }

    // Handle Rename Symbol mode
    if this.mode == Mode::RenameSymbol {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_rename_symbol(cx);
                return;
            }
            "enter" => {
                this.confirm_rename_symbol(cx);
                return;
            }
            "backspace" => {
                this.rename_symbol_backspace(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for rename
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars().filter(|c| !c.is_control()) {
                    this.rename_symbol_insert_char(c, cx);
                }
                return;
            }
        }
    }

    // Handle Create Named Range mode
    if this.mode == Mode::CreateNamedRange {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_create_named_range(cx);
                return;
            }
            "enter" => {
                this.confirm_create_named_range(cx);
                return;
            }
            "backspace" => {
                this.create_name_backspace(cx);
                return;
            }
            "tab" => {
                this.create_name_tab(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for create name
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars().filter(|c| !c.is_control()) {
                    this.create_name_insert_char(c, cx);
                }
                return;
            }
        }
        return; // Consume all keystrokes in create named range mode
    }

    // Handle Edit Description mode
    if this.mode == Mode::EditDescription {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_edit_description(cx);
                return;
            }
            "enter" => {
                this.apply_edit_description(cx);
                return;
            }
            "backspace" => {
                this.edit_description_backspace(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for description
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars().filter(|c| !c.is_control()) {
                    this.edit_description_insert_char(c, cx);
                }
                return;
            }
        }
    }

    // Handle Tour mode
    if this.mode == Mode::Tour {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_tour(cx);
                return;
            }
            "enter" | "right" => {
                if this.tour_step < 3 {
                    this.tour_next(cx);
                } else {
                    this.tour_done(cx);
                }
                return;
            }
            "left" => {
                this.tour_back(cx);
                return;
            }
            _ => {}
        }
        return; // Consume all keystrokes in tour mode
    }

    // Handle Impact Preview mode
    if this.mode == Mode::ImpactPreview {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_impact_preview(cx);
                return;
            }
            "enter" => {
                this.apply_impact_preview(cx);
                return;
            }
            _ => {}
        }
        return; // Consume all keystrokes in impact preview mode
    }

    // Handle Refactor Log mode
    if this.mode == Mode::RefactorLog {
        match event.keystroke.key.as_str() {
            "escape" | "enter" => {
                this.hide_refactor_log(cx);
                return;
            }
            _ => {}
        }
        return; // Consume all keystrokes in refactor log mode
    }

    // Handle Extract Named Range mode
    if this.mode == Mode::ExtractNamedRange {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_extract_named_range(cx);
                return;
            }
            "enter" => {
                this.confirm_extract_named_range(cx);
                return;
            }
            "backspace" => {
                match this.extract_focus {
                    CreateNameFocus::Name => this.extract_name_backspace(cx),
                    CreateNameFocus::Description => this.extract_description_backspace(cx),
                }
                return;
            }
            "tab" => {
                this.extract_tab(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for extract name/description
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars().filter(|c| !c.is_control()) {
                    match this.extract_focus {
                        CreateNameFocus::Name => this.extract_name_insert_char(c, cx),
                        CreateNameFocus::Description => this.extract_description_insert_char(c, cx),
                    }
                }
                return;
            }
        }
        return; // Consume all keystrokes in extract mode
    }

    // Handle License mode
    if this.mode == Mode::License {
        match event.keystroke.key.as_str() {
            "escape" => {
                this.hide_license(cx);
                return;
            }
            "enter" => {
                if !this.license_input.is_empty() {
                    this.apply_license(cx);
                }
                return;
            }
            "backspace" => {
                this.license_backspace(cx);
                return;
            }
            _ => {}
        }

        // Handle text input for license
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars().filter(|c| !c.is_control()) {
                    this.license_insert_char(c, cx);
                }
                return;
            }
        }
        return; // Consume all keystrokes in license mode
    }

    // Handle About mode (consume keystrokes, only escape closes)
    if this.mode == Mode::About {
        if event.keystroke.key == "escape" {
            this.hide_about(cx);
        }
        return; // Consume all keystrokes in about mode
    }

    // Handle Preferences mode (consume keystrokes, only escape closes)
    if this.mode == Mode::Preferences {
        if event.keystroke.key == "escape" {
            this.hide_preferences(cx);
        }
        return; // Consume all keystrokes in preferences mode
    }

    // Handle Hint mode (Vimium-style jump navigation)
    if this.mode == Mode::Hint {
        let key = event.keystroke.key.as_str();
        if this.apply_hint_key(key, cx) {
            return;
        }
        // Unhandled key in hint mode - exit without action
        this.exit_hint_mode(cx);
        return;
    }

    // Handle Navigation mode with keyboard hints or vim mode enabled
    // (before Names tab filter and before regular text input)
    if this.mode == Mode::Navigation
        && !event.keystroke.modifiers.control
        && !event.keystroke.modifiers.alt
        && !event.keystroke.modifiers.platform
        && !event.keystroke.modifiers.shift
    {
        let key = event.keystroke.key.as_str();
        let hints_enabled = this.keyboard_hints_enabled(cx);
        let vim_enabled = this.vim_mode_enabled(cx);

        // 'g' enters command/hint mode (if hints OR vim mode enabled)
        // - With hints: shows cell labels + g-commands (gg)
        // - With vim only: just g-commands (gg), no labels
        if (hints_enabled || vim_enabled) && key == "g" {
            this.enter_hint_mode_with_labels(hints_enabled, cx);
            return;
        }

        // Vim keys (if vim mode enabled)
        if vim_enabled {
            if this.apply_vim_key(key, cx) {
                return;
            }
            // For non-vim keys in vim mode, don't start edit - just ignore
            if key.len() == 1 && key.chars().next().map(|c| c.is_alphabetic()).unwrap_or(false) {
                // In vim mode, typing letters doesn't start edit
                // Show status hint for unknown vim key
                this.status_message = Some(format!("Unknown vim key: {}", key));
                cx.notify();
                return;
            }
        }
    }

    // Handle Names tab filter input (when inspector visible + Names tab + Navigation mode)
    if this.inspector_visible
        && this.inspector_tab == InspectorTab::Names
        && this.mode == Mode::Navigation
    {
        match event.keystroke.key.as_str() {
            "escape" => {
                if !this.names_filter_query.is_empty() {
                    // Clear filter first
                    this.names_filter_query.clear();
                    cx.notify();
                } else {
                    // Exit preview when closing inspector
                    if this.is_previewing() {
                        this.exit_preview(cx);
                    }
                    // Close inspector
                    this.inspector_visible = false;
                    cx.notify();
                }
                return;
            }
            "backspace" => {
                if !this.names_filter_query.is_empty() {
                    this.names_filter_query.pop();
                    cx.notify();
                }
                return;
            }
            "/" => {
                // "/" focuses filter (already focused, this just prevents "/" from being typed)
                return;
            }
            _ => {}
        }

        // Handle text input for filter (only alphanumeric and underscore)
        if let Some(key_char) = &event.keystroke.key_char {
            if !event.keystroke.modifiers.control
                && !event.keystroke.modifiers.alt
                && !event.keystroke.modifiers.platform
            {
                for c in key_char.chars().filter(|c| c.is_alphanumeric() || *c == '_' || *c == '.') {
                    this.names_filter_query.push(c);
                }
                cx.notify();
                return;
            }
        }
    }

    // Handle GoTo mode
    if this.mode == Mode::GoTo {
        if event.keystroke.key == "enter" {
            this.confirm_goto(cx);
            return;
        } else if event.keystroke.key == "escape" {
            this.hide_goto(cx);
            return;
        } else if event.keystroke.key == "backspace" {
            this.goto_backspace(cx);
            return;
        }
    }

    // Handle Find mode
    if this.mode == Mode::Find {
        if event.keystroke.key == "escape" {
            this.hide_find(cx);
            return;
        } else if event.keystroke.key == "backspace" {
            this.find_backspace(cx);
            return;
        } else if event.keystroke.key == "tab" {
            // Tab toggles focus between find and replace inputs
            this.find_toggle_focus(cx);
            return;
        }
    }

    // Handle HubPasteToken mode
    if this.mode == Mode::HubPasteToken {
        if event.keystroke.key == "escape" {
            this.hub_cancel_sign_in(cx);
            return;
        } else if event.keystroke.key == "enter" {
            this.hub_complete_sign_in(cx);
            return;
        } else if event.keystroke.key == "backspace" {
            this.hub_token_backspace(cx);
            return;
        }
    }

    // Handle HubLink mode
    if this.mode == Mode::HubLink {
        if event.keystroke.key == "escape" {
            this.hub_cancel_link(cx);
            return;
        } else if event.keystroke.key == "backspace" {
            this.hub_dataset_backspace(cx);
            return;
        }
    }

    if let Some(key_char) = &event.keystroke.key_char {
        if !event.keystroke.modifiers.control
            && !event.keystroke.modifiers.alt
            && !event.keystroke.modifiers.platform
        {
            // Filter out control characters - let them be handled by actions instead
            let printable_chars: String = key_char.chars()
                .filter(|c| !c.is_control())
                .collect();

            if !printable_chars.is_empty() {
                match this.mode {
                    Mode::GoTo => {
                        for c in printable_chars.chars() {
                            this.goto_insert_char(c, cx);
                        }
                    }
                    Mode::Find => {
                        for c in printable_chars.chars() {
                            this.find_insert_char(c, cx);
                        }
                    }
                    Mode::HubPasteToken => {
                        for c in printable_chars.chars() {
                            this.hub_token_insert_char(c, cx);
                        }
                    }
                    Mode::HubLink => {
                        for c in printable_chars.chars() {
                            this.hub_dataset_insert_char(c, cx);
                        }
                    }
                    _ => {
                        for c in printable_chars.chars() {
                            this.insert_char(c, cx);
                        }
                        this.update_edit_scroll(window);
                    }
                }
            }
        }
    }
}
