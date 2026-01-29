mod about_dialog;
mod ai_settings_dialog;
mod ask_ai_dialog;
mod color_picker;
pub mod command_palette;
mod hub_dialogs;
mod export_report_dialog;
mod filter_dropdown;
mod find_dialog;
mod font_picker;
mod formula_bar;
mod goto_dialog;
mod grid;
mod headers;
pub mod impact_preview;
mod import_overlay;
mod import_report_dialog;
pub mod inspector_panel;
mod keytips_overlay;
#[cfg(feature = "pro")]
mod lua_console;
#[cfg(not(feature = "pro"))]
mod lua_console_stub;
#[cfg(not(feature = "pro"))]
use lua_console_stub as lua_console;
pub mod license_dialog;
mod paste_special_dialog;
mod preferences_panel;
pub mod refactor_log;
mod menu_bar;
mod status_bar;
mod theme_picker;
mod tour;
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
    let show_keytips = app.keytips_active;
    let show_rewind_confirm = app.rewind_confirm.visible;
    let show_rewind_success = app.rewind_success.visible;
    let show_import_overlay = app.import_overlay_visible;
    let show_name_tooltip = app.should_show_name_tooltip(cx) && app.mode == Mode::Navigation;
    let show_f2_tip = app.should_show_f2_tip(cx);  // Show immediately on trigger, not gated on mode
    let show_inspector = app.inspector_visible;
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
            if event.keystroke.key == "f1" && this.f1_help_visible {
                this.f1_help_visible = false;
                cx.notify();
            }
            if event.keystroke.key == "space" && this.is_previewing() {
                this.exit_preview(cx);
            }
        }))
        // Mouse wheel scrolling (or zoom with Ctrl/Cmd)
        .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, _, cx| {
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
                let new_height = (this.lua_console.resize_start_height + delta)
                    .max(crate::scripting::MIN_CONSOLE_HEIGHT)
                    .min(crate::scripting::MAX_CONSOLE_HEIGHT);
                this.lua_console.height = new_height;
                cx.notify();
                return; // Don't process other drags while resizing console
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

            // Chrome scrim: subtle gradient fade from titlebar into content
            let scrim_top = titlebar_bg.opacity(0.12);
            let scrim_bottom = titlebar_bg.opacity(0.0);

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
            // Chrome scrim: subtle top fade for visual separation
            .child(
                div()
                    .w_full()
                    .h(px(8.0))  // Slightly smaller scrim
                    .flex_shrink_0()
                    .bg(linear_gradient(
                        180.0,
                        linear_color_stop(scrim_top, 0.0),
                        linear_color_stop(scrim_bottom, 1.0),
                    ))
            )
        })
        // Hide in-app menu bar on macOS (uses native menu bar instead)
        // Also hide in zen mode
        .when(!cfg!(target_os = "macos") && !zen_mode, |d| {
            d.child(menu_bar::render_menu_bar(app, cx))
        })
        .when(!zen_mode, |div| {
            div.child(formula_bar::render_formula_bar(app, window, cx))
        })
        .child(headers::render_column_headers(app, cx))
        // Split view: render two grids side-by-side, or single grid
        .child(if app.is_split() {
            render_split_grids(app, window, cx).into_any_element()
        } else {
            grid::render_grid(app, window, cx, None).into_any_element()
        })
        // Lua console panel (above status bar)
        .child(lua_console::render_lua_console(app, cx))
        .when(!zen_mode, |div| {
            div.child(status_bar::render_status_bar(app, editing, cx))
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
        .when(show_rewind_confirm, |div| {
            div.child(rewind_dialogs::render_rewind_confirm_dialog(app, cx))
        })
        .when(show_rewind_success, |div| {
            div.child(rewind_dialogs::render_rewind_success_banner(app, cx))
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
        // NOTE: Autocomplete, signature help, and error banner popups are now rendered
        // in the grid overlay layer (grid.rs::render_popup_overlay) where they can be
        // positioned relative to the cell rect without menu/formula bar offset math.
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
