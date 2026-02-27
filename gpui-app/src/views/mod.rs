mod about_dialog;
mod ai_settings_dialog;
mod ask_ai_dialog;
mod color_picker;
pub mod command_palette;
pub(crate) mod context_menu;
mod cycle_banner;
mod hub_dialogs;
mod export_report_dialog;
mod filter_dropdown;
mod find_dialog;
mod font_picker;
pub(crate) mod format_bar;
mod formula_bar;
mod goto_dialog;
mod grid;
mod headers;
pub mod impact_preview;
mod import_overlay;
mod import_report_dialog;
pub mod inspector_panel;
pub mod profiler_panel;
mod keytips_overlay;
mod code_render;
pub(crate) mod lua_console;
pub(crate) mod script_view;
pub(crate) mod terminal_panel;
pub mod license_dialog;
pub mod minimap;
mod paste_special_dialog;
mod convert_picker;
mod preferences_panel;
pub mod refactor_log;
mod menu_bar;
mod status_bar;
mod theme_picker;
mod tour;
mod number_format_dialog;
mod validation_dialog;
mod validation_dropdown_view;

// Extracted modules
mod actions_nav;
mod actions_edit;
mod actions_ui;
mod key_handler;
mod f1_help;
mod named_range_dialogs;
mod rewind_dialogs;
mod transform_diff_dialog;

use gpui::*;
use gpui::prelude::FluentBuilder;
use crate::app::{Spreadsheet, CELL_HEIGHT};
use crate::mode::Mode;
use crate::theme::TokenKey;

pub fn render_spreadsheet(app: &mut Spreadsheet, window: &mut Window, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    // Check if validation dropdown source has changed (fingerprint mismatch)
    app.check_dropdown_staleness(cx);

    // Auto-dismiss rewind success banner after 5 seconds
    if app.rewind_success.should_dismiss() {
        app.rewind_success.hide();
    }

    let editing = app.mode.is_editing();
    let show_goto = app.mode == Mode::GoTo;
    let show_find = app.mode == Mode::Find;
    let show_command = app.mode == Mode::Command;
    let show_font_picker = app.mode == Mode::FontPicker;
    let show_theme_picker = app.mode == Mode::ThemePicker;
    let show_about = app.mode == Mode::About;
    let show_rename_symbol = app.mode == Mode::RenameSymbol;
    let show_create_named_range = app.mode == Mode::CreateNamedRange;
    let show_edit_description = app.mode == Mode::EditDescription;
    let show_tour = app.mode == Mode::Tour;
    let show_impact_preview = app.mode == Mode::ImpactPreview;
    let show_refactor_log = app.mode == Mode::RefactorLog;
    let show_extract_named_range = app.mode == Mode::ExtractNamedRange;
    let show_import_report = app.mode == Mode::ImportReport;
    let show_export_report = app.mode == Mode::ExportReport;
    let show_preferences = app.mode == Mode::Preferences;
    let show_license = app.mode == Mode::License;
    let show_hub_paste_token = app.mode == Mode::HubPasteToken;
    let show_hub_link = app.mode == Mode::HubLink;
    let show_hub_publish_confirm = app.mode == Mode::HubPublishConfirm;
    let show_validation_dialog = app.mode == Mode::ValidationDialog;
    let show_ai_settings = app.mode == Mode::AISettings;
    let show_ask_ai = app.mode == Mode::AiDialog;
    let show_explain_diff = app.mode == Mode::ExplainDiff;
    let show_paste_special = app.mode == Mode::PasteSpecial;
    let show_color_picker = app.mode == Mode::ColorPicker;
    let show_number_format_editor = app.mode == Mode::NumberFormatEditor;
    let show_transform_preview = app.mode == Mode::TransformPreview;
    let show_convert_picker = app.mode == Mode::ConvertPicker;
    let show_keytips = app.keytips_active;
    let show_rewind_confirm = app.rewind_confirm.visible;
    let show_rewind_success = app.rewind_success.visible;
    let show_cycle_banner = app.cycle_banner.visible;
    let show_merge_confirm = app.merge_confirm.visible;
    let show_close_confirm = app.close_confirm_visible;
    let show_approval_confirm = app.approval_confirm_visible;
    let show_approval_drift = app.approval_drift_visible;
    let show_import_overlay = app.import_overlay_visible;
    let show_name_tooltip = app.should_show_name_tooltip(cx) && app.mode == Mode::Navigation;
    let show_f2_tip = app.should_show_f2_tip(cx);  // Show immediately on trigger, not gated on mode
    let show_inspector = app.inspector_visible;
    let show_profiler = app.profiler_visible;
    let zen_mode = app.zen_mode;

    // Build element with action handlers from extracted modules
    let el = div()
        .relative()
        .key_context("Spreadsheet")
        .track_focus(&app.focus_handle);
    let el = actions_nav::bind(el, cx);
    let el = actions_edit::bind(el, cx);
    let el = actions_ui::bind(el, cx);

    el
        // Character input (handles editing, goto, find, and command modes)
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            key_handler::handle_key_down(this, event, window, cx);
        }))
        // Key release handling:
        // - F1 hold-to-peek: hide help when F1 is released
        // - Space hold-to-peek: exit preview when Space is released
        // - Option double-tap: KeyTips detection (macOS only)
        .on_key_up(cx.listener(|this, event: &KeyUpEvent, _, cx| {
            // Clear focus invariant marker on key_up to prevent stale state
            // across event boundaries (IME, key repeat, composition).
            #[cfg(debug_assertions)]
            crate::views::terminal_panel::clear_terminal_key_handled();

            if event.keystroke.key == "f1" && this.f1_help_visible {
                this.f1_help_visible = false;
                cx.notify();
            }
            if event.keystroke.key == "space" && this.is_previewing() {
                this.exit_preview(cx);
            }
        }))
        // Mouse wheel scrolling (or zoom with Ctrl/Cmd)
        .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, window, cx| {
            // Don't scroll the grid when the terminal panel has focus —
            // the terminal's own scroll handler handles it and calls stop_propagation,
            // but this guard is defense-in-depth for edge cases.
            if this.terminal_has_focus(window) {
                return;
            }
            // Check for zoom modifier (Ctrl on Linux/Windows, Cmd on macOS)
            // macOS: use .platform (Cmd), others: use .control (Ctrl)
            #[cfg(target_os = "macos")]
            let zoom_modifier = event.modifiers.platform;
            #[cfg(not(target_os = "macos"))]
            let zoom_modifier = event.modifiers.control;

            let delta = event.delta.pixel_delta(px(CELL_HEIGHT));
            let dy: f32 = delta.y.into();

            if zoom_modifier {
                // Zoom: Ctrl/Cmd + wheel
                this.zoom_wheel(dy, cx);
                return; // Don't scroll
            }

            // Normal scrolling
            let dx: f32 = delta.x.into();
            let delta_rows = (-dy / CELL_HEIGHT).round() as i32;
            let delta_cols = (-dx / CELL_HEIGHT).round() as i32;
            if delta_rows != 0 || delta_cols != 0 {
                this.scroll(delta_rows, delta_cols, cx);
            }
        }))
        // Mouse move for resize and header selection dragging
        .on_mouse_move(cx.listener(|this, event: &MouseMoveEvent, _, cx| {
            // Handle Lua console resize drag
            if this.lua_console.resizing {
                let y: f32 = event.position.y.into();
                let delta = this.lua_console.resize_start_y - y; // Inverted: dragging up increases height
                let new_height = this.lua_console.resize_start_height + delta;
                this.lua_console.set_height_from_drag(new_height);
                cx.notify();
                return; // Don't process other drags while resizing console
            }
            // Handle Terminal resize drag
            if this.terminal.resizing {
                let y: f32 = event.position.y.into();
                let delta = this.terminal.resize_start_y - y; // Inverted: dragging up increases height
                let new_height = this.terminal.resize_start_height + delta;
                this.terminal.set_height_from_drag(new_height);
                cx.notify();
                return;
            }
            // Handle column resize drag
            if let Some(col) = this.resizing_col {
                let x: f32 = event.position.x.into();
                let delta = x - this.resize_start_pos;
                let new_width = this.resize_start_size + delta;
                this.set_col_width(col, new_width);
                cx.notify();
            }
            // Handle row resize drag
            if let Some(row) = this.resizing_row {
                let y: f32 = event.position.y.into();
                let delta = y - this.resize_start_pos;
                let new_height = this.resize_start_size + delta;
                this.set_row_height(row, new_height);
                cx.notify();
            }
            // Handle column header selection drag
            if this.dragging_col_header {
                let x: f32 = event.position.x.into();
                if let Some(col) = this.col_from_window_x(x) {
                    this.continue_col_header_drag(col, cx);
                }
            }
            // Handle row header selection drag
            if this.dragging_row_header {
                let y: f32 = event.position.y.into();
                if let Some(row) = this.row_from_window_y(y) {
                    this.continue_row_header_drag(row, cx);
                }
            }
        }))
        // Mouse up to end resize and header selection drag
        .on_mouse_up(MouseButton::Left, cx.listener(|this, _, _, cx| {
            // End Lua console resize
            if this.lua_console.resizing {
                this.lua_console.resizing = false;
                cx.notify();
            }
            // End Terminal resize
            if this.terminal.resizing {
                this.terminal.resizing = false;
                cx.notify();
            }
            // End column/row resize and record to history (coalescing)
            if let Some(col) = this.resizing_col.take() {
                // Record the sizing change to history (coalesces all drag events into one entry)
                let old = this.resize_start_original.take();
                this.record_col_width_change(col, old, cx);
                cx.notify();
            }
            if let Some(row) = this.resizing_row.take() {
                // Record the sizing change to history (coalesces all drag events into one entry)
                let old = this.resize_start_original.take();
                this.record_row_height_change(row, old, cx);
                cx.notify();
            }
            // End header selection drag
            if this.dragging_row_header {
                this.end_row_header_drag(cx);
            }
            if this.dragging_col_header {
                this.end_col_header_drag(cx);
            }
        }))
        .flex()
        .flex_col()
        .size_full()
        .bg(app.token(TokenKey::AppBg))
        // On macOS with transparent titlebar, render a draggable title bar area
        // This allows window dragging and double-click to zoom
        .when(cfg!(target_os = "macos"), |d| {
            let titlebar_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);

            // Typography: Zed-style approach
            // Primary text: default text color (let macOS handle inactive dimming)
            // Secondary text: muted + slight fade for extra restraint
            let primary_color = app.token(TokenKey::TextPrimary);
            let secondary_color = app.token(TokenKey::TextMuted).opacity(0.85);

            let title_primary = app.document_meta.title_primary(app.history.is_dirty());
            let title_secondary = app.document_meta.title_secondary();


            // Check if we should show the default app prompt
            // Also check timer for success state auto-hide
            app.check_default_app_prompt_timer(cx);

            let show_default_prompt = app.should_show_default_app_prompt(cx);
            let prompt_state = app.default_app_prompt_state;
            let prompt_file_type = app.get_prompt_file_type();

            // Mark shown when transitioning from Hidden to visible
            // This records timestamp for cool-down and prevents session spam
            if show_default_prompt && prompt_state == crate::app::DefaultAppPromptState::Hidden {
                app.on_default_app_prompt_shown(cx);
            }

            d.child(
                div()
                    .id("macos-title-bar")
                    .w_full()
                    .h(px(34.0))  // Match Zed's titlebar height
                    .flex_shrink_0()
                    .bg(titlebar_bg)
                    .border_b_1()
                    .border_color(panel_border.opacity(0.5))  // Subtle hairline
                    .window_control_area(WindowControlArea::Drag)
                    // Double-click to zoom (macOS native behavior)
                    .on_click(|event, window, _cx| {
                        if event.click_count() == 2 {
                            window.titlebar_double_click();
                        }
                    })
                    // Left padding for traffic lights
                    .pl(px(72.0))
                    .pr(px(12.0))  // Right padding for prompt chip
                    .flex()
                    .items_center()
                    .justify_between()  // Push left content and right prompt apart
                    // Left side: document title
                    .child(
                        div()
                            .flex()
                            .items_center()
                            .gap(px(6.0))
                            // Primary: filename + dirty indicator
                            .child(
                                div()
                                    .text_color(primary_color)
                                    .text_size(px(12.0))  // Smaller, calmer
                                    .child(title_primary)
                            )
                            // Secondary: provenance (quieter, extra small)
                            .when_some(title_secondary, |el, secondary| {
                                el.child(
                                    div()
                                        .text_color(secondary_color)
                                        .text_size(px(10.0))  // XSmall like Zed
                                        .child(secondary)
                                )
                            })
                    )
                    // Right side: default app prompt chip (state-based rendering)
                    .when(show_default_prompt, |el| {
                        use crate::app::DefaultAppPromptState;

                        let border_color = panel_border.opacity(0.2);
                        let muted = app.token(TokenKey::TextMuted);
                        let chip_bg = titlebar_bg.opacity(0.5);  // Muted background like inactive tab

                        // File type name for scoped messaging
                        let file_type_name = prompt_file_type
                            .map(|ft| ft.short_name())
                            .unwrap_or(".csv files");

                        match prompt_state {
                            // Success state: brief confirmation then auto-hide
                            DefaultAppPromptState::Success => {
                                el.child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .px(px(8.0))
                                        .py(px(3.0))
                                        .rounded_sm()
                                        .bg(chip_bg)
                                        .border_1()
                                        .border_color(border_color)
                                        .child(
                                            div()
                                                .text_color(muted)
                                                .text_size(px(10.0))
                                                .child("Default set")
                                        )
                                )
                            }

                            // Needs Settings state: user must complete in System Settings
                            DefaultAppPromptState::NeedsSettings => {
                                el.child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .rounded_sm()
                                        .bg(chip_bg)
                                        .border_1()
                                        .border_color(border_color)
                                        .child(
                                            div()
                                                .id("default-app-open-settings")
                                                .flex()
                                                .items_center()
                                                .gap(px(6.0))
                                                .px(px(8.0))
                                                .py(px(3.0))
                                                .cursor_pointer()
                                                .hover(|s| s.bg(panel_border.opacity(0.08)))
                                                .on_click(cx.listener(|this, _, _, cx| {
                                                    this.open_default_app_settings(cx);
                                                }))
                                                .child(
                                                    div()
                                                        .text_color(muted)
                                                        .text_size(px(10.0))
                                                        .child("Finish in System Settings")
                                                )
                                                .child(
                                                    div()
                                                        .text_color(muted.opacity(0.8))
                                                        .text_size(px(10.0))
                                                        .child("Open")
                                                )
                                        )
                                )
                            }

                            // Normal showing state: prompt to set default
                            _ => {
                                el.child(
                                    div()
                                        .flex()
                                        .items_center()
                                        .rounded_sm()
                                        .bg(chip_bg)
                                        .border_1()
                                        .border_color(border_color)
                                        // Main clickable area
                                        .child(
                                            div()
                                                .id("default-app-set")
                                                .flex()
                                                .items_center()
                                                .gap(px(6.0))
                                                .px(px(8.0))
                                                .py(px(3.0))
                                                .cursor_pointer()
                                                .hover(|s| s.bg(panel_border.opacity(0.08)))
                                                .on_click(cx.listener(|this, _, _, cx| {
                                                    this.set_as_default_app(cx);
                                                }))
                                                // Label: "Open .csv files with VisiGrid"
                                                .child(
                                                    div()
                                                        .text_color(muted.opacity(0.8))
                                                        .text_size(px(10.0))
                                                        .child(format!("Open {} with VisiGrid", file_type_name))
                                                )
                                                // Button: subtle outline style
                                                .child(
                                                    div()
                                                        .text_color(muted)
                                                        .text_size(px(10.0))
                                                        .child("Make default")
                                                )
                                        )
                                        // Close button (dismiss forever)
                                        .child(
                                            div()
                                                .border_l_1()
                                                .border_color(border_color)
                                                .child(
                                                    div()
                                                        .id("default-app-dismiss")
                                                        .px(px(6.0))
                                                        .py(px(3.0))
                                                        .cursor_pointer()
                                                        .text_color(muted.opacity(0.4))  // Low contrast until hover
                                                        .text_size(px(10.0))
                                                        .hover(|s| s.bg(panel_border.opacity(0.08)).text_color(muted.opacity(0.8)))
                                                        .on_click(cx.listener(|this, _, _, cx| {
                                                            this.dismiss_default_app_prompt(cx);
                                                        }))
                                                        .child("\u{2715}")  // ✕
                                                )
                                        )
                                )
                            }
                        }
                    })
            )
            // No scrim - let formula bar background do the work
        })
        // Hide in-app menu bar on macOS (uses native menu bar instead)
        // Also hide in zen mode
        .when(!cfg!(target_os = "macos") && !zen_mode, |d| {
            d.child(menu_bar::render_menu_bar(app, cx))
        })
        .when(!zen_mode, |div| {
            div.child(formula_bar::render_formula_bar(app, window, cx))
        })
        // Inline formula helper strip (signature hint, shown only when editing a formula)
        .when_some(if zen_mode { None } else { app.signature_help() }, |div, sig_info| {
            let panel_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);
            let text_primary = app.token(TokenKey::TextPrimary);
            let text_muted = app.token(TokenKey::TextMuted);
            let accent = app.token(TokenKey::Accent);
            div.child(formula_bar::render_formula_helper_strip(
                &sig_info,
                panel_bg,
                panel_border,
                text_primary,
                text_muted,
                accent,
            ))
        })
        .when(!zen_mode && {
            use crate::settings::{Setting, user_settings};
            match &user_settings(cx).appearance.show_format_bar {
                Setting::Value(v) => *v,
                Setting::Inherit => true,
            }
        }, |div| {
            div.child(format_bar::render_format_bar(app, window, cx))
        })
        .child(headers::render_column_headers(app, cx))
        // Split view: render two grids side-by-side, or single grid
        // Wrapped in flex-row to accommodate optional minimap strip on the right
        .child({
            if app.script.open {
                script_view::render_script_view(app, window, cx).into_any_element()
            } else {
                let show_minimap = app.minimap_visible && !zen_mode;
                let grid_element = if app.is_split() {
                    render_split_grids(app, window, cx).into_any_element()
                } else {
                    grid::render_grid(app, window, cx, None).into_any_element()
                };
                div()
                    .flex()
                    .flex_row()
                    .flex_1()
                    .min_h(px(0.0))  // Allow grid to shrink below content size for console panel
                    .child(grid_element)
                    .when(show_minimap, |d| {
                        d.child(minimap::render_minimap(app, window, cx))
                    })
                    .into_any_element()
            }
        })
        // Bottom panel: tabbed container for Lua console + Terminal
        .child({
            // Always pump Lua debug events even when panel is hidden
            lua_console::pump_debug_events(app, cx);
            lua_console::refresh_input_tokens(app);
            render_bottom_panel(app, window, cx)
        })
        .when(!zen_mode, |div| {
            div.child(status_bar::render_status_bar(app, editing, cx))
        })
        // Debug: layout origin overlay (Cmd+Alt+Shift+G / Ctrl+Alt+Shift+G).
        // Draws grid_body_origin line and row header resize-band geometry
        // in window coordinates. If the cyan line doesn't sit exactly at the
        // top of the first data row, top_chrome_height() is wrong.
        .when(app.debug_grid_alignment, |d| {
            let origin_y = app.grid_layout.grid_body_origin.1;
            let origin_x = app.grid_layout.grid_body_origin.0;
            let row_h = app.metrics.cell_h;
            let visible = app.visible_rows().min(20); // cap overlay to avoid perf hit
            let grab = crate::app::ROW_RESIZE_GRAB_PX;

            let mut overlay = d
                // Horizontal cyan line at grid_body_origin.y
                .child(
                    div().absolute()
                        .left_0()
                        .top(px(origin_y))
                        .w_full()
                        .h(px(1.0))
                        .bg(gpui::rgba(0x00ffffff))
                )
                // Vertical cyan line at grid_body_origin.x
                .child(
                    div().absolute()
                        .left(px(origin_x))
                        .top(px(origin_y))
                        .w(px(1.0))
                        .h(px(app.grid_layout.viewport_size.1))
                        .bg(gpui::rgba(0x00ffffff))
                );

            // Row header rects + resize bands for visible rows
            for i in 0..visible {
                let row = app.view_state.scroll_row + i;
                let y = origin_y + app.row_y_offset(row);
                let h = app.metrics.row_height(app.row_height(row));

                // Row header outline (yellow, 1px)
                overlay = overlay.child(
                    div().absolute()
                        .left_0()
                        .top(px(y))
                        .w(px(origin_x))
                        .h(px(h))
                        .border_1()
                        .border_color(gpui::rgba(0xffff0060))
                );
                // Resize grab band (red translucent, bottom ROW_RESIZE_GRAB_PX)
                overlay = overlay.child(
                    div().absolute()
                        .left_0()
                        .top(px(y + h - grab))
                        .w(px(origin_x))
                        .h(px(grab))
                        .bg(gpui::rgba(0xff000040))
                );
            }
            overlay
        })
        // Font size dropdown overlay — rendered at root level so it paints above
        // column headers and grid cells. Only visible when dropdown is open.
        .when(app.ui.format_bar.size_dropdown, |d| {
            d.child(format_bar::render_font_size_dropdown(app, cx))
        })
        // Format dropdown overlay (Bold/Italic/Underline/Alignment)
        .when(app.ui.format_menu_open, |d| {
            d.child(format_bar::render_format_dropdown(app, cx))
        })
        // Number format quick-menu dropdown (123 ▾ button)
        .when(app.ui.format_bar.number_format_menu_open, |d| {
            d.child(format_bar::render_number_format_dropdown(app, cx))
        })
        // Cell styles quick-menu dropdown (Styles ▾ button)
        .when(app.ui.format_bar.cell_style_menu_open, |d| {
            d.child(format_bar::render_cell_style_dropdown(app, cx))
        })
        // Inspector panel (right-side drawer) with click-outside-to-close backdrop.
        // Rendered BEFORE modal overlays so modals sit on top in z-order.
        .when(show_inspector, |d| {
            d.child(
                gpui::div()
                    .id("inspector-backdrop")
                    .absolute()
                    .inset_0()
                    // Transparent background makes the element "hit testable" - without this,
                    // GPUI may pass events through to elements behind it
                    .bg(hsla(0.0, 0.0, 0.0, 0.0))
                    // Click backdrop to close inspector (outside the panel)
                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                        // Don't close inspector when a modal overlay is open -
                        // the modal should handle the click, not us
                        if this.mode.is_overlay() {
                            return;
                        }
                        // Stop propagation first to prevent grid from receiving this event
                        cx.stop_propagation();
                        this.inspector_visible = false;
                        cx.notify();
                    }))
                    // Also stop mouse up to prevent any grid selection from completing
                    .on_mouse_up(MouseButton::Left, |_, _, cx| {
                        cx.stop_propagation();
                    })
                    // The panel itself stops propagation so clicks on it don't close
                    .child(inspector_panel::render_inspector_panel(app, cx))
            )
        })
        // Profiler panel (right-side drawer, mutually exclusive with inspector)
        .when(show_profiler, |d| {
            d.child(
                gpui::div()
                    .id("profiler-backdrop")
                    .absolute()
                    .inset_0()
                    .bg(hsla(0.0, 0.0, 0.0, 0.0))
                    .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
                        if this.mode.is_overlay() { return; }
                        cx.stop_propagation();
                        this.profiler_visible = false;
                        window.focus(&this.focus_handle, cx);
                        cx.notify();
                    }))
                    .on_mouse_up(MouseButton::Left, |_, _, cx| {
                        cx.stop_propagation();
                    })
                    .child(profiler_panel::render_profiler_panel(app, cx))
            )
        })
        .when(show_goto, |div| {
            div.child(goto_dialog::render_goto_dialog(app))
        })
        .when(show_find, |div| {
            div.child(find_dialog::render_find_dialog(app))
        })
        .when(show_command, |div| {
            div.child(command_palette::render_command_palette(app, cx))
        })
        .when(show_font_picker, |div| {
            div.child(font_picker::render_font_picker(app, cx))
        })
        .when(show_color_picker, |div| {
            div.child(color_picker::render_color_picker(app, cx))
        })
        .when(show_theme_picker, |div| {
            div.child(theme_picker::render_theme_picker(app, cx))
        })
        .when(show_preferences, |div| {
            div.child(preferences_panel::render_preferences_panel(app, cx))
        })
        .when(show_about, |div| {
            div.child(about_dialog::render_about_dialog(app, cx))
        })
        .when(show_paste_special, |div| {
            div.child(paste_special_dialog::render_paste_special_dialog(app, cx))
        })
        .when(show_convert_picker, |div| {
            div.child(convert_picker::render_convert_picker(app, cx))
        })
        .when(show_transform_preview, |div| {
            div.child(transform_diff_dialog::render_transform_diff_dialog(app, cx))
        })
        .when(show_number_format_editor, |div| {
            div.child(number_format_dialog::render_number_format_dialog(app, cx))
        })
        // KeyTips overlay (macOS Option double-tap accelerators)
        .when(show_keytips, |div| {
            div.child(keytips_overlay::render_keytips_overlay(app, cx))
        })
        .when(show_license, |div| {
            div.child(license_dialog::render_license_dialog(app, cx))
        })
        .when(show_validation_dialog, |div| {
            div.child(validation_dialog::render_validation_dialog(app, cx))
        })
        .when(show_ai_settings, |div| {
            div.child(ai_settings_dialog::render_ai_settings_dialog(app, cx))
        })
        .when(show_ask_ai, |div| {
            div.child(ask_ai_dialog::render_ask_ai_dialog(app, cx))
        })
        // Ask AI context menu (rendered as overlay above the dialog)
        .when_some(ask_ai_dialog::render_ask_ai_context_menu(app, cx), |div, menu| {
            div.child(menu)
        })
        .when(show_explain_diff, |div| {
            div.child(inspector_panel::render_explain_diff_dialog(app, cx))
        })
        // History entry context menu (right-click menu)
        .when_some(inspector_panel::render_history_context_menu(app, cx), |div, menu| {
            div.child(menu)
        })
        // Cell/header right-click context menu
        .when_some(context_menu::render_context_menu(app, window, cx), |div, menu| {
            div.child(menu)
        })
        .when(show_rewind_confirm, |div| {
            div.child(rewind_dialogs::render_rewind_confirm_dialog(app, cx))
        })
        .when(show_merge_confirm, |div| {
            div.child(render_merge_confirm_dialog(app, cx))
        })
        .when(show_close_confirm, |div| {
            div.child(render_close_confirm_dialog(app, window, cx))
        })
        .when(show_approval_confirm, |div| {
            div.child(render_approval_confirm_dialog(app, cx))
        })
        .when(show_approval_drift, |div| {
            div.child(render_approval_drift_panel(app, cx))
        })
        .when(show_rewind_success, |div| {
            div.child(rewind_dialogs::render_rewind_success_banner(app, cx))
        })
        .when(show_cycle_banner, |div| {
            div.child(cycle_banner::render_cycle_banner(app, cx))
        })
        .when(show_hub_paste_token, |div| {
            div.child(hub_dialogs::render_paste_token_dialog(app, cx))
        })
        .when(show_hub_link, |div| {
            div.child(hub_dialogs::render_link_dialog(app, cx))
        })
        .when(show_hub_publish_confirm, |div| {
            div.child(hub_dialogs::render_publish_confirm_dialog(app, cx))
        })
        .when(show_import_report, |div| {
            div.child(import_report_dialog::render_import_report_dialog(app, cx))
        })
        .when(show_export_report, |div| {
            div.child(export_report_dialog::render_export_report_dialog(app, cx))
        })
        // Import overlay: shows during background Excel imports (after 150ms delay)
        .when(show_import_overlay, |div| {
            div.child(import_overlay::render_import_overlay(app, cx))
        })
        .when(show_rename_symbol, |div| {
            div.child(named_range_dialogs::render_rename_symbol_dialog(app))
        })
        .when(show_create_named_range, |div| {
            div.child(named_range_dialogs::render_create_named_range_dialog(app, cx))
        })
        .when(show_edit_description, |div| {
            div.child(named_range_dialogs::render_edit_description_dialog(app))
        })
        .when(show_tour, |div| {
            div.child(tour::render_tour(app, cx))
        })
        .when(show_impact_preview, |div| {
            div.child(impact_preview::render_impact_preview(app, cx))
        })
        .when(show_refactor_log, |div| {
            div.child(refactor_log::render_refactor_log(app, cx))
        })
        .when(show_extract_named_range, |div| {
            div.child(named_range_dialogs::render_extract_named_range_dialog(app))
        })
        // Name tooltip (one-time first-run hint)
        .when(show_name_tooltip, |div| {
            div.child(tour::render_name_tooltip(app, cx))
        })
        // F2 function key tip (macOS only, shown when editing via non-F2 path)
        .when(show_f2_tip, |div| {
            div.child(tour::render_f2_tooltip(app, cx))
        })
        // Menu dropdown (only on non-macOS where we have in-app menu)
        .when(!cfg!(target_os = "macos") && app.open_menu.is_some(), |div| {
            div.child(menu_bar::render_menu_dropdown(app, cx))
        })
        // Filter dropdown popup
        .when_some(filter_dropdown::render_filter_dropdown(app, cx), |div, dropdown| {
            div.child(dropdown)
        })
        // Validation dropdown popup (list validation)
        .when_some(validation_dropdown_view::render_validation_dropdown(app, cx), |div, dropdown| {
            div.child(dropdown)
        })
        // Inspector panel was moved above modal overlays (rendered after status bar)
        // so that modals sit on top in z-order and receive clicks first.
        // NOTE: Autocomplete and error banner popups are rendered in the grid overlay
        // layer (grid.rs::render_popup_overlay). Signature help is now an inline strip
        // rendered in the main layout flow between formula bar and format bar.
        // Hover documentation popup (when not editing and hovering over formula bar)
        .when_some(app.hover_function.filter(|_| !app.mode.is_editing() && !app.autocomplete_visible), |div, func| {
            let panel_bg = app.token(TokenKey::PanelBg);
            let panel_border = app.token(TokenKey::PanelBorder);
            let text_primary = app.token(TokenKey::TextPrimary);
            let text_muted = app.token(TokenKey::TextMuted);
            let accent = app.token(TokenKey::Accent);
            div.child(formula_bar::render_hover_docs(func, panel_bg, panel_border, text_primary, text_muted, accent))
        })
        // F1 hold-to-peek context help overlay
        .when_some(
            app.f1_help_visible.then(|| f1_help::render_f1_help_overlay(app, cx)),
            |div, overlay| div.child(overlay)
        )
}

/// Render the merge cells data-loss confirmation dialog.
fn render_merge_confirm_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let warning_color = app.token(TokenKey::Error);

    let affected = &app.merge_confirm.affected_cells;
    let data_msg = if affected.len() > 10 {
        format!("Data in {} cells will be lost.", affected.len())
    } else {
        format!("Data in {} will be lost.", affected.join(", "))
    };

    div()
        .absolute()
        .inset_0()
        .bg(hsla(0.0, 0.0, 0.0, 0.6))
        .flex()
        .items_center()
        .justify_center()
        .child(
            div()
                .id("merge-confirm-dialog")
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .w(px(380.0))
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                    if event.keystroke.key == "escape" {
                        this.merge_confirm.visible = false;
                        cx.notify();
                    }
                }))
                // Header
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(20.0))
                                .text_color(warning_color)
                                .child("\u{26A0}")
                        )
                        .child(
                            div()
                                .text_size(px(16.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child("Merge Cells")
                        )
                )
                // Body
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(13.0))
                                .text_color(text_primary)
                                .child("Merging cells only keeps the upper-left value and discards other values.")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child(SharedString::from(data_msg))
                        )
                )
                // Footer buttons
                .child(
                    div()
                        .flex()
                        .justify_end()
                        .gap_2()
                        .child(
                            div()
                                .id("merge-cancel-btn")
                                .px_3()
                                .py_1()
                                .border_1()
                                .border_color(panel_border)
                                .rounded_sm()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .cursor_pointer()
                                .hover(|s| s.text_color(text_primary))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.merge_confirm.visible = false;
                                    cx.notify();
                                }))
                                .child("Cancel")
                        )
                        .child(
                            div()
                                .id("merge-confirm-btn")
                                .px_3()
                                .py_1()
                                .bg(accent)
                                .rounded_sm()
                                .text_size(px(12.0))
                                .text_color(rgb(0xffffff))
                                .cursor_pointer()
                                .hover(|s| s.bg(accent.opacity(0.85)))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.merge_confirm.visible = false;
                                    this.merge_cells_confirmed(cx);
                                }))
                                .child("Merge Anyway")
                        )
                )
        )
}

/// Render the close-window save confirmation dialog.
fn render_close_confirm_dialog(app: &Spreadsheet, _window: &mut Window, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    use crate::ui::{modal_backdrop, Button, DialogFrame, DialogSize};

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let focused = app.close_confirm_focused;

    let filename = app.current_file.as_ref()
        .and_then(|p| p.file_name())
        .and_then(|n| n.to_str())
        .unwrap_or(&app.document_meta.display_name);

    let body_msg = format!("Do you want to save changes to \"{}\"?", filename);

    // Body content
    let body = div()
        .flex()
        .flex_col()
        .gap_2()
        .child(
            div()
                .text_size(px(13.0))
                .text_color(text_primary)
                .child(SharedString::from(body_msg))
        )
        .child(
            div()
                .text_size(px(12.0))
                .text_color(text_muted)
                .child("Your changes will be lost if you don't save.")
        );

    // Footer buttons with focus ring
    let cancel_btn = Button::new("close-cancel-btn", "Cancel")
        .secondary(if focused == 0 { accent } else { panel_border }, if focused == 0 { text_primary } else { text_muted })
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.close_confirm_visible = false;
            cx.notify();
        }));

    let dont_save_btn = Button::new("close-dont-save-btn", "Don't Save")
        .secondary(if focused == 1 { accent } else { panel_border }, if focused == 1 { text_primary } else { text_muted })
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            this.close_confirm_visible = false;
            this.prepare_close(cx);
            window.remove_window();
        }));

    let save_btn = Button::new("close-save-btn", "Save")
        .primary(accent, rgb(0xffffff).into())
        .when(focused == 2, |b| b.border_1().border_color(hsla(0.0, 0.0, 1.0, 1.0)))
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            this.close_confirm_visible = false;
            let saved = this.save_and_close(cx);
            if saved {
                this.prepare_close(cx);
                window.remove_window();
            }
        }));

    let footer = div()
        .flex()
        .justify_end()
        .gap_2()
        .child(cancel_btn)
        .child(dont_save_btn)
        .child(save_btn);

    let header = div()
        .text_size(px(14.0))
        .font_weight(FontWeight::SEMIBOLD)
        .text_color(text_primary)
        .child("Save Changes");

    modal_backdrop(
        "close-confirm-dialog",
        DialogFrame::new(body, panel_bg, panel_border)
            .size(DialogSize::Md)
            .header(header)
            .footer(footer),
    )
}

/// Render the approval confirmation dialog (when re-approving after drift)
fn render_approval_confirm_dialog(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let warning_color = app.token(TokenKey::Warn);
    let error_color = app.token(TokenKey::Error);

    let label_input = app.approval_label_input.clone();
    let (action_count, cell_count) = app.approval_drift_count();
    let show_count_warning = action_count > 1 || cell_count > 10;

    div()
        .absolute()
        .inset_0()
        .bg(hsla(0.0, 0.0, 0.0, 0.6))
        .flex()
        .items_center()
        .justify_center()
        .child(
            div()
                .id("approval-confirm-dialog")
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .w(px(400.0))
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                    if event.keystroke.key == "escape" {
                        this.cancel_approval_confirm(cx);
                    } else if event.keystroke.key == "enter" {
                        let label = if this.approval_label_input.is_empty() {
                            None
                        } else {
                            Some(this.approval_label_input.clone())
                        };
                        this.approve_model_confirmed(label, cx);
                    } else if let Some(key_char) = &event.keystroke.key_char {
                        if !event.keystroke.modifiers.control && !event.keystroke.modifiers.platform {
                            this.approval_label_input.push_str(key_char);
                            cx.notify();
                        }
                    } else if event.keystroke.key == "backspace" {
                        this.approval_label_input.pop();
                        cx.notify();
                    }
                }))
                // Header
                .child(
                    div()
                        .flex()
                        .items_center()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(20.0))
                                .text_color(warning_color)
                                .child("⚠")
                        )
                        .child(
                            div()
                                .text_size(px(16.0))
                                .font_weight(FontWeight::SEMIBOLD)
                                .text_color(text_primary)
                                .child("Approve new logic?")
                        )
                )
                // Body
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        .child(
                            div()
                                .text_size(px(13.0))
                                .text_color(text_primary)
                                .child("This will replace the previously approved logic.")
                        )
                        .child(
                            div()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .child("Any prior verification will no longer apply.")
                        )
                        // Change count warning (when significant)
                        .when(show_count_warning, |d| {
                            let msg = if cell_count > 0 {
                                format!("⚠ {} logic change(s) affecting {} cell(s) detected", action_count, cell_count)
                            } else {
                                format!("⚠ {} logic change(s) detected", action_count)
                            };
                            d.child(
                                div()
                                    .text_size(px(12.0))
                                    .text_color(error_color)
                                    .font_weight(FontWeight::MEDIUM)
                                    .child(SharedString::from(msg))
                            )
                        })
                )
                // Label input
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_1()
                        .child(
                            div()
                                .text_size(px(11.0))
                                .text_color(text_muted)
                                .child("Label (optional):")
                        )
                        .child(
                            div()
                                .px_2()
                                .py_1()
                                .bg(panel_bg)
                                .border_1()
                                .border_color(panel_border)
                                .rounded_sm()
                                .text_size(px(12.0))
                                .text_color(text_primary)
                                .min_h(px(24.0))
                                .child(if label_input.is_empty() {
                                    SharedString::from("e.g., Q3 Final, Reviewed, v2.1")
                                } else {
                                    SharedString::from(label_input)
                                })
                                .when(app.approval_label_input.is_empty(), |d| {
                                    d.text_color(text_muted.opacity(0.5))
                                })
                        )
                )
                // Footer buttons
                .child(
                    div()
                        .flex()
                        .justify_end()
                        .gap_2()
                        .child(
                            div()
                                .id("approval-cancel-btn")
                                .px_3()
                                .py_1()
                                .border_1()
                                .border_color(panel_border)
                                .rounded_sm()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .cursor_pointer()
                                .hover(|s| s.text_color(text_primary))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.cancel_approval_confirm(cx);
                                }))
                                .child("Cancel")
                        )
                        .child(
                            div()
                                .id("approval-confirm-btn")
                                .px_3()
                                .py_1()
                                .bg(accent)
                                .rounded_sm()
                                .text_size(px(12.0))
                                .text_color(rgb(0xffffff))
                                .cursor_pointer()
                                .hover(|s| s.bg(accent.opacity(0.85)))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    let label = if this.approval_label_input.is_empty() {
                                        None
                                    } else {
                                        Some(this.approval_label_input.clone())
                                    };
                                    this.approve_model_confirmed(label, cx);
                                }))
                                .child("Approve New Logic")
                        )
                )
        )
}

/// Render the "Why drifted?" panel showing changes since approval
fn render_approval_drift_panel(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let warning_color = app.token(TokenKey::Warn);
    let error_color = app.token(TokenKey::Error);
    let ok_color = app.token(TokenKey::Ok);

    let changes = app.approval_drift_changes();
    let fp_comparison = app.semantic_fingerprint_comparison(cx);

    div()
        .absolute()
        .inset_0()
        .bg(hsla(0.0, 0.0, 0.0, 0.6))
        .flex()
        .items_center()
        .justify_center()
        .child(
            div()
                .id("approval-drift-panel")
                .bg(panel_bg)
                .border_1()
                .border_color(panel_border)
                .rounded_md()
                .shadow_lg()
                .w(px(500.0))
                .max_h(px(400.0))
                .p_4()
                .flex()
                .flex_col()
                .gap_3()
                .on_key_down(cx.listener(|this, event: &KeyDownEvent, _, cx| {
                    if event.keystroke.key == "escape" {
                        this.hide_approval_drift(cx);
                    }
                }))
                // Header
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_between()
                        .child(
                            div()
                                .flex()
                                .items_center()
                                .gap_2()
                                .child(
                                    div()
                                        .text_size(px(20.0))
                                        .text_color(warning_color)
                                        .child("⚠")
                                )
                                .child(
                                    div()
                                        .text_size(px(16.0))
                                        .font_weight(FontWeight::SEMIBOLD)
                                        .text_color(text_primary)
                                        .child("Changes since approval")
                                )
                        )
                        .child(
                            div()
                                .id("drift-close-btn")
                                .px_2()
                                .py_1()
                                .cursor_pointer()
                                .text_color(text_muted)
                                .hover(|s| s.text_color(text_primary))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_approval_drift(cx);
                                }))
                                .child("✕")
                        )
                )
                // Body - list of changes
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_2()
                        .overflow_hidden()
                        .max_h(px(250.0))
                        .children(
                            if changes.is_empty() {
                                // No history changes - show fingerprint comparison if available
                                if let Some((expected, current)) = fp_comparison.clone() {
                                    if expected != current {
                                        vec![
                                            div()
                                                .flex()
                                                .flex_col()
                                                .gap_2()
                                                .child(
                                                    div()
                                                        .text_size(px(13.0))
                                                        .text_color(text_muted)
                                                        .child("File was modified outside this session.")
                                                )
                                                .child(
                                                    div()
                                                        .flex()
                                                        .flex_col()
                                                        .gap_1()
                                                        .p_2()
                                                        .bg(panel_border.opacity(0.3))
                                                        .rounded_sm()
                                                        .text_size(px(12.0))
                                                        .font_family("monospace")
                                                        .child(
                                                            div()
                                                                .flex()
                                                                .gap_2()
                                                                .child(div().text_color(text_muted).child("Expected:"))
                                                                .child(div().text_color(error_color).child(expected))
                                                        )
                                                        .child(
                                                            div()
                                                                .flex()
                                                                .gap_2()
                                                                .child(div().text_color(text_muted).child("Current: "))
                                                                .child(div().text_color(ok_color).child(current))
                                                        )
                                                )
                                                .into_any_element()
                                        ]
                                    } else {
                                        vec![
                                            div()
                                                .text_size(px(13.0))
                                                .text_color(text_muted)
                                                .child("No semantic changes detected.")
                                                .into_any_element()
                                        ]
                                    }
                                } else {
                                    vec![
                                        div()
                                            .text_size(px(13.0))
                                            .text_color(text_muted)
                                            .child("No semantic changes detected.")
                                            .into_any_element()
                                    ]
                                }
                            } else {
                                changes.into_iter().map(|(label, location, cells)| {
                                    let label: SharedString = label.into();
                                    let location_str: Option<SharedString> = location.map(|s| s.into());

                                    let mut row = div()
                                        .flex()
                                        .flex_col()
                                        .gap_1()
                                        .p_2()
                                        .bg(panel_border.opacity(0.3))
                                        .rounded_sm()
                                        // Action label
                                        .child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .gap_2()
                                                .child(
                                                    div()
                                                        .text_size(px(13.0))
                                                        .font_weight(FontWeight::MEDIUM)
                                                        .text_color(text_primary)
                                                        .child(label)
                                                )
                                                .when_some(location_str, |d, loc| {
                                                    d.child(
                                                        div()
                                                            .text_size(px(11.0))
                                                            .text_color(text_muted)
                                                            .child(loc)
                                                    )
                                                })
                                        );

                                    // Cell changes
                                    for (addr, old, new) in cells {
                                        let addr: SharedString = addr.into();
                                        let old_display: SharedString = if old.is_empty() { "(empty)".into() } else { old.into() };
                                        let new_display: SharedString = if new.is_empty() { "(empty)".into() } else { new.into() };

                                        row = row.child(
                                            div()
                                                .flex()
                                                .items_center()
                                                .gap_2()
                                                .text_size(px(11.0))
                                                .child(
                                                    div()
                                                        .text_color(accent)
                                                        .font_weight(FontWeight::MEDIUM)
                                                        .child(addr)
                                                )
                                                .child(
                                                    div()
                                                        .text_color(error_color.opacity(0.8))
                                                        .child(old_display)
                                                )
                                                .child(
                                                    div()
                                                        .text_color(text_muted)
                                                        .child("→")
                                                )
                                                .child(
                                                    div()
                                                        .text_color(ok_color)
                                                        .child(new_display)
                                                )
                                        );
                                    }

                                    row.into_any_element()
                                }).collect()
                            }
                        )
                )
                // Info: what counts as semantics
                .child(
                    div()
                        .flex()
                        .flex_col()
                        .gap_1()
                        .p_2()
                        .bg(panel_border.opacity(0.2))
                        .rounded_sm()
                        .text_size(px(10.0))
                        .text_color(text_muted)
                        .child(
                            div()
                                .font_weight(FontWeight::MEDIUM)
                                .child("Approval tracks logic, not appearance:")
                        )
                        .child("✓ Formulas, values, references, structure")
                        .child("✗ Formatting, colors, fonts, column widths")
                )
                // Footer buttons
                .child(
                    div()
                        .flex()
                        .justify_end()
                        .gap_2()
                        .child(
                            div()
                                .id("drift-dismiss-btn")
                                .px_3()
                                .py_1()
                                .border_1()
                                .border_color(panel_border)
                                .rounded_sm()
                                .text_size(px(12.0))
                                .text_color(text_muted)
                                .cursor_pointer()
                                .hover(|s| s.text_color(text_primary))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_approval_drift(cx);
                                }))
                                .child("Close")
                        )
                        .child(
                            div()
                                .id("drift-approve-btn")
                                .px_3()
                                .py_1()
                                .bg(accent)
                                .rounded_sm()
                                .text_size(px(12.0))
                                .text_color(rgb(0xffffff))
                                .cursor_pointer()
                                .hover(|s| s.bg(accent.opacity(0.85)))
                                .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                                    this.hide_approval_drift(cx);
                                    this.approve_model(None, cx);
                                }))
                                .child("Approve New Logic")
                        )
                )
        )
}

/// Render the bottom panel with tab bar (Lua console / Terminal).
///
/// Both panels share a single bottom panel area. The tab bar lets users switch
/// between them. Each shortcut (`Alt+F11` for Lua, `Ctrl+`` for Terminal) opens
/// its tab or toggles the panel closed if that tab is already active.
fn render_bottom_panel(
    app: &Spreadsheet,
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    use crate::app::BottomPanelTab;

    if !app.bottom_panel_visible {
        return div().into_any_element();
    }

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let active_tab = app.bottom_panel_tab;

    // Tab bar
    let tab_bar = div()
        .flex()
        .flex_row()
        .items_center()
        .gap(px(2.0))
        .px(px(8.0))
        .h(px(26.0))
        .bg(panel_bg)
        .border_b_1()
        .border_color(panel_border)
        .child(render_panel_tab(
            "bottom-tab-lua",
            "Lua",
            active_tab == BottomPanelTab::Lua,
            accent, text_primary, text_muted, panel_border,
            cx.listener(|this, _, window, cx| {
                use crate::app::BottomPanelTab;
                this.bottom_panel_tab = BottomPanelTab::Lua;
                this.lua_console.visible = true;
                this.terminal.visible = false;
                this.terminal_focused = false;
                if this.lua_console.first_open {
                    this.lua_console.show();
                }
                window.focus(&this.console_focus_handle, cx);
                cx.notify();
            }),
        ))
        .child(render_panel_tab(
            "bottom-tab-terminal",
            "Terminal",
            active_tab == BottomPanelTab::Terminal,
            accent, text_primary, text_muted, panel_border,
            cx.listener(|this, _, window, cx| {
                use crate::app::BottomPanelTab;
                this.bottom_panel_tab = BottomPanelTab::Terminal;
                this.lua_console.visible = false;
                this.terminal.visible = true;
                this.terminal_focused = true;
                if this.terminal.term.is_none() && !this.terminal.exited {
                    this.spawn_terminal(window, cx);
                } else {
                    this.terminal.ensure_cwd();
                }
                window.focus(&this.terminal_focus_handle, cx);
                cx.notify();
            }),
        ));

    // Panel content (only one is visible at a time)
    let content = match active_tab {
        BottomPanelTab::Lua => {
            lua_console::render_lua_console(app, cx).into_any_element()
        }
        BottomPanelTab::Terminal => {
            terminal_panel::render_terminal_panel(app, window, cx).into_any_element()
        }
    };

    div()
        .flex()
        .flex_col()
        .flex_shrink_0()
        .child(tab_bar)
        .child(content)
        .into_any_element()
}

/// Render a single tab button for the bottom panel tab bar.
fn render_panel_tab(
    id: &'static str,
    label: &'static str,
    is_active: bool,
    accent: Hsla,
    text_primary: Hsla,
    text_muted: Hsla,
    border_color: Hsla,
    on_click: impl Fn(&ClickEvent, &mut Window, &mut App) + 'static,
) -> impl IntoElement {
    div()
        .id(id)
        .cursor_pointer()
        .px(px(10.0))
        .py(px(4.0))
        .text_size(px(11.0))
        .rounded_t(px(3.0))
        .when(is_active, |d| {
            d.text_color(text_primary)
                .font_weight(FontWeight::SEMIBOLD)
                .border_b_2()
                .border_color(accent)
        })
        .when(!is_active, |d| {
            d.text_color(text_muted)
                .hover(|s| s.text_color(text_primary))
        })
        .child(label)
        .on_click(on_click)
}

/// Render split view with two grids side-by-side (50/50)
fn render_split_grids(
    app: &mut Spreadsheet,
    window: &Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    use crate::split_view::SplitSide;

    let accent = app.token(TokenKey::Accent);
    let panel_border = app.token(TokenKey::PanelBorder);
    let active_side = app.split_active_side;

    // Border width for active pane indicator
    let active_border_width = px(2.0);

    div()
        .flex()
        .flex_row()
        .size_full()
        // Left pane (uses main view_state)
        .child(
            div()
                .flex_1()
                .h_full()
                .overflow_hidden()
                .border_r_1()
                .border_color(panel_border)
                .when(active_side == SplitSide::Left, |d| {
                    d.border_2().border_color(accent)
                })
                .child(grid::render_grid(app, window, cx, Some(SplitSide::Left)))
        )
        // Vertical divider (implicit via border)
        // Right pane (uses split_pane.view_state for independent scroll/selection)
        .child(
            div()
                .flex_1()
                .h_full()
                .overflow_hidden()
                .when(active_side == SplitSide::Right, |d| {
                    d.border_2().border_color(accent)
                })
                .child(grid::render_grid(app, window, cx, Some(SplitSide::Right)))
        )
}
