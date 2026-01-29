use gpui::*;
use crate::app::Spreadsheet;
use crate::actions::*;
use crate::mode::InspectorTab;
use crate::formatting::BorderApplyMode;
use crate::search::MenuCategory;

pub(crate) fn bind(
    el: Div,
    cx: &mut Context<Spreadsheet>,
) -> Div {
    el
        // File actions
        // Note: NewWindow (Ctrl+N) is handled at App level in main.rs
        // NewInPlace replaces current workbook (dangerous, not bound by default)
        .on_action(cx.listener(|this, _: &NewInPlace, window, cx| {
            this.new_in_place(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenFile, _, cx| {
            this.open_file(cx);
        }))
        .on_action(cx.listener(|this, _: &Save, _, cx| {
            this.save(cx);
        }))
        .on_action(cx.listener(|this, _: &SaveAs, _, cx| {
            this.save_as(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportCsv, _, cx| {
            this.export_csv(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportTsv, _, cx| {
            this.export_tsv(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportJson, _, cx| {
            this.export_json(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportXlsx, _, cx| {
            this.export_xlsx(cx);
        }))
        .on_action(cx.listener(|this, _: &ExportProvenance, _, cx| {
            this.export_provenance(cx);
        }))
        // VisiHub sync actions
        .on_action(cx.listener(|this, _: &HubCheckStatus, _, cx| {
            this.hub_check_status(cx);
        }))
        .on_action(cx.listener(|this, _: &HubPull, _, cx| {
            this.hub_pull(cx);
        }))
        .on_action(cx.listener(|this, _: &HubOpenRemoteAsCopy, _, cx| {
            this.hub_open_remote_as_copy(cx);
        }))
        .on_action(cx.listener(|this, _: &HubUnlink, _, cx| {
            this.hub_unlink(cx);
        }))
        .on_action(cx.listener(|this, _: &HubDiagnostics, _, cx| {
            this.hub_diagnostics(cx);
        }))
        .on_action(cx.listener(|this, _: &HubSignIn, _, cx| {
            this.hub_sign_in(cx);
        }))
        .on_action(cx.listener(|this, _: &HubSignOut, _, cx| {
            this.hub_sign_out(cx);
        }))
        .on_action(cx.listener(|this, _: &HubLink, _, cx| {
            this.hub_show_link_dialog(cx);
        }))
        .on_action(cx.listener(|this, _: &HubPublish, _, cx| {
            this.hub_publish(cx);
        }))
        // Data operations (sort/filter)
        .on_action(cx.listener(|this, _: &SortAscending, window, cx| {
            this.sort_by_current_column(visigrid_engine::filter::SortDirection::Ascending, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &SortDescending, window, cx| {
            this.sort_by_current_column(visigrid_engine::filter::SortDirection::Descending, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleAutoFilter, _, cx| {
            this.toggle_auto_filter(cx);
        }))
        .on_action(cx.listener(|this, _: &ClearSort, window, cx| {
            this.clear_sort(cx);
            this.update_title_if_needed(window, cx);
        }))
        // Data validation
        .on_action(cx.listener(|this, _: &ShowDataValidation, _, cx| {
            // TODO: Show data validation dialog
            this.status_message = Some("Data validation dialog not yet implemented".to_string());
            cx.notify();
        }))
        // Insert Formula with AI (Ctrl+Shift+A)
        .on_action(cx.listener(|this, _: &InsertFormula, _, cx| {
            this.show_ask_ai(cx);
        }))
        // Analyze with AI (Ctrl+Shift+E)
        .on_action(cx.listener(|this, _: &Analyze, _, cx| {
            this.show_analyze(cx);
        }))
        .on_action(cx.listener(|this, _: &OpenValidationDropdown, _, cx| {
            this.open_validation_dropdown(cx);
        }))
        .on_action(cx.listener(|this, _: &AutoSum, _, cx| {
            this.autosum(cx);
        }))
        .on_action(cx.listener(|this, _: &TrimWhitespace, _, cx| {
            this.trim_whitespace(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleFormulaView, _, cx| {
            this.toggle_show_formulas(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleShowZeros, _, cx| {
            this.toggle_show_zeros(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleInspector, _, cx| {
            this.inspector_visible = !this.inspector_visible;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &ToggleZenMode, _, cx| {
            this.zen_mode = !this.zen_mode;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &ToggleVerifiedMode, _, cx| {
            this.toggle_verified_mode(cx);
        }))
        .on_action(cx.listener(|this, _: &Recalculate, _, cx| {
            this.recalculate(cx);
        }))
        .on_action(cx.listener(|this, _: &NavPerfReport, _, cx| {
            let msg = this.nav_perf.report()
                .unwrap_or_else(|| "Nav perf tracking disabled. Set VISIGRID_PERF=nav and restart.".into());
            this.status_message = Some(msg);
            cx.notify();
        }))
        // Zoom
        .on_action(cx.listener(|this, _: &ZoomIn, _, cx| {
            this.zoom_in(cx);
        }))
        .on_action(cx.listener(|this, _: &ZoomOut, _, cx| {
            this.zoom_out(cx);
        }))
        .on_action(cx.listener(|this, _: &ZoomReset, _, cx| {
            this.zoom_reset(cx);
        }))
        // Freeze panes
        .on_action(cx.listener(|this, _: &FreezeTopRow, _, cx| {
            this.freeze_top_row(cx);
        }))
        .on_action(cx.listener(|this, _: &FreezeFirstColumn, _, cx| {
            this.freeze_first_column(cx);
        }))
        .on_action(cx.listener(|this, _: &FreezePanes, _, cx| {
            this.freeze_panes(cx);
        }))
        .on_action(cx.listener(|this, _: &UnfreezePanes, _, cx| {
            this.unfreeze_panes(cx);
        }))
        // Split view
        .on_action(cx.listener(|this, _: &SplitRight, _, cx| {
            this.split_right(cx);
        }))
        .on_action(cx.listener(|this, _: &CloseSplit, _, cx| {
            this.close_split(cx);
        }))
        .on_action(cx.listener(|this, _: &FocusOtherPane, _, cx| {
            this.focus_other_pane(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleTrace, _, cx| {
            this.toggle_trace(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleDebugGridAlignment, _, cx| {
            this.debug_grid_alignment = !this.debug_grid_alignment;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &CycleTracePrecedent, window, cx| {
            // Block while validation dropdown or any modal is open
            if this.is_validation_dropdown_open() || this.mode.is_overlay() {
                return;
            }
            let reverse = window.modifiers().shift;
            this.cycle_trace_precedent(reverse, cx);
        }))
        .on_action(cx.listener(|this, _: &CycleTraceDependent, window, cx| {
            // Block while validation dropdown or any modal is open
            if this.is_validation_dropdown_open() || this.mode.is_overlay() {
                return;
            }
            let reverse = window.modifiers().shift;
            this.cycle_trace_dependent(reverse, cx);
        }))
        .on_action(cx.listener(|this, _: &ReturnToTraceSource, _, cx| {
            // Block while validation dropdown or any modal is open
            if this.is_validation_dropdown_open() || this.mode.is_overlay() {
                return;
            }
            this.return_to_trace_source(cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleLuaConsole, window, cx| {
            // Pro feature gate
            if !visigrid_license::is_feature_enabled("lua") {
                this.status_message = Some("Lua scripting requires VisiGrid Pro".to_string());
                cx.notify();
                return;
            }
            this.lua_console.toggle();
            if this.lua_console.visible {
                window.focus(&this.console_focus_handle, cx);
            }
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &ShowFormatPanel, _, cx| {
            this.inspector_visible = true;
            this.inspector_tab = InspectorTab::Format;
            cx.notify();
        }))
        .on_action(cx.listener(|this, _: &ShowHistoryPanel, _, cx| {
            this.inspector_visible = true;
            this.inspector_tab = InspectorTab::History;
            cx.notify();
        }))
        // Formatting
        .on_action(cx.listener(|this, _: &ToggleBold, window, cx| {
            this.toggle_bold(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleItalic, window, cx| {
            this.toggle_italic(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleUnderline, window, cx| {
            this.toggle_underline(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &ToggleStrikethrough, window, cx| {
            this.toggle_strikethrough(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &FormatCurrency, window, cx| {
            this.format_currency(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &FormatPercent, window, cx| {
            this.format_percent(cx);
            this.update_title_if_needed(window, cx);
        }))
        // Background colors
        .on_action(cx.listener(|this, _: &ClearBackground, window, cx| {
            this.set_background_color(None, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &ClearFormatting, window, cx| {
            this.clear_formatting_selection(cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &FormatPainter, _, cx| {
            this.start_format_painter(cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundYellow, window, cx| {
            this.set_background_color(Some([255, 255, 0, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundGreen, window, cx| {
            this.set_background_color(Some([198, 239, 206, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundBlue, window, cx| {
            this.set_background_color(Some([189, 215, 238, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundRed, window, cx| {
            this.set_background_color(Some([255, 199, 206, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundOrange, window, cx| {
            this.set_background_color(Some([255, 235, 156, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundPurple, window, cx| {
            this.set_background_color(Some([204, 192, 218, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundGray, window, cx| {
            this.set_background_color(Some([217, 217, 217, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BackgroundCyan, window, cx| {
            this.set_background_color(Some([183, 222, 232, 255]), cx);
            this.update_title_if_needed(window, cx);
        }))
        // Borders
        .on_action(cx.listener(|this, _: &BordersAll, window, cx| {
            this.apply_borders(BorderApplyMode::All, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BordersOutline, window, cx| {
            this.apply_borders(BorderApplyMode::Outline, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BordersInside, window, cx| {
            this.apply_borders(BorderApplyMode::Inside, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BordersTop, window, cx| {
            this.apply_borders(BorderApplyMode::Top, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BordersBottom, window, cx| {
            this.apply_borders(BorderApplyMode::Bottom, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BordersLeft, window, cx| {
            this.apply_borders(BorderApplyMode::Left, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BordersRight, window, cx| {
            this.apply_borders(BorderApplyMode::Right, cx);
            this.update_title_if_needed(window, cx);
        }))
        .on_action(cx.listener(|this, _: &BordersClear, window, cx| {
            this.apply_borders(BorderApplyMode::Clear, cx);
            this.update_title_if_needed(window, cx);
        }))
        // Command palette
        .on_action(cx.listener(|this, _: &ToggleCommandPalette, _, cx| {
            this.toggle_palette(cx);
        }))
        .on_action(cx.listener(|this, _: &QuickOpen, _, cx| {
            if this.mode.is_editing() || this.mode.is_formula() {
                return;  // Don't steal Ctrl+K during cell editing
            }
            this.show_quick_open(cx);
        }))
        // View actions (for native macOS menus)
        .on_action(cx.listener(|this, _: &ShowAbout, _, cx| {
            this.show_about(cx);
        }))
        .on_action(cx.listener(|this, _: &ShowLicense, _, cx| {
            this.show_license(cx);
        }))
        .on_action(cx.listener(|this, _: &ShowFontPicker, window, cx| {
            this.show_font_picker(window, cx);
        }))
        .on_action(cx.listener(|this, _: &ShowColorPicker, window, cx| {
            this.show_color_picker(crate::color_palette::ColorTarget::Fill, window, cx);
        }))
        .on_action(cx.listener(|this, _: &ShowKeyTips, _, cx| {
            this.toggle_keytips(cx);
        }))
        .on_action(cx.listener(|this, _: &ShowPreferences, _, cx| {
            this.show_preferences(cx);
        }))
        .on_action(cx.listener(|this, _: &OpenKeybindings, _, cx| {
            this.open_keybindings(cx);
        }))
        // Window menu actions
        .on_action(cx.listener(|_this, _: &Minimize, window, _cx| {
            window.minimize_window();
        }))
        .on_action(cx.listener(|_this, _: &Zoom, window, _cx| {
            window.zoom_window();
        }))
        .on_action(cx.listener(|_this, _: &BringAllToFront, _window, cx| {
            // Activate the app to bring all windows to front
            cx.activate(true);
        }))
        .on_action(cx.listener(|this, _: &CloseWindow, window, cx| {
            // Commit any pending edit before closing
            this.commit_pending_edit(cx);

            // If workbook is clean, just close
            if !this.is_modified {
                // Unregister from window registry before closing
                this.unregister_from_window_registry(cx);
                window.remove_window();
                return;
            }

            // Workbook is dirty - prompt user to save
            let filename = this.current_file.as_ref()
                .and_then(|p| p.file_name())
                .and_then(|n| n.to_str())
                .unwrap_or("Untitled");

            let receiver = window.prompt(
                gpui::PromptLevel::Warning,
                &format!("Do you want to save changes to \"{}\"?", filename),
                Some("Your changes will be lost if you don't save them."),
                &["Save", "Don't Save", "Cancel"],
                cx,
            );

            // Capture window handle for closing from async context
            let window_handle = window.window_handle();

            cx.spawn(async move |this, cx| {
                if let Ok(response) = receiver.await {
                    match response {
                        0 => {
                            // Save, then close
                            let save_succeeded = this.update(cx, |this, cx| {
                                this.save_and_close(cx)
                            }).unwrap_or(false);

                            if save_succeeded {
                                // Unregister from window registry before closing
                                let _ = this.update(cx, |this, cx| {
                                    this.unregister_from_window_registry(cx);
                                });
                                let _ = window_handle.update(cx, |_, window, _| {
                                    window.remove_window();
                                });
                            }
                        }
                        1 => {
                            // Don't Save - close without saving
                            // Unregister from window registry before closing
                            let _ = this.update(cx, |this, cx| {
                                this.unregister_from_window_registry(cx);
                            });
                            let _ = window_handle.update(cx, |_, window, _| {
                                window.remove_window();
                            });
                        }
                        _ => {
                            // Cancel - do nothing
                        }
                    }
                }
            }).detach();
        }))
        .on_action(cx.listener(|this, _: &Quit, _, cx| {
            // Commit any pending edit before quitting
            this.commit_pending_edit(cx);
            // Propagate to global quit handler (saves session and quits)
            cx.propagate();
        }))
        // Menu bar (Alt+letter accelerators)
        .on_action(cx.listener(|this, _: &OpenFileMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::File, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenEditMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Edit, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenViewMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::View, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenInsertMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Insert, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenFormatMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Format, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenDataMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Data, cx);
        }))
        .on_action(cx.listener(|this, _: &OpenHelpMenu, _, cx| {
            this.toggle_menu(crate::mode::Menu::Help, cx);
        }))
        // Alt accelerators (open Command Palette scoped to menu category)
        // Guarded by should_handle_option_accelerators() to avoid conflicts with
        // macOS character composition (accents, special characters) when typing.
        .on_action(cx.listener(|this, _: &AltFile, _, cx| {
            if this.should_handle_option_accelerators() {
                this.apply_menu_scope(MenuCategory::File, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltEdit, _, cx| {
            if this.should_handle_option_accelerators() {
                this.apply_menu_scope(MenuCategory::Edit, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltView, _, cx| {
            if this.should_handle_option_accelerators() {
                this.apply_menu_scope(MenuCategory::View, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltFormat, _, cx| {
            if this.should_handle_option_accelerators() {
                this.apply_menu_scope(MenuCategory::Format, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltData, _, cx| {
            if this.should_handle_option_accelerators() {
                this.apply_menu_scope(MenuCategory::Data, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltTools, _, cx| {
            if this.should_handle_option_accelerators() {
                this.apply_menu_scope(MenuCategory::Tools, cx);
            }
        }))
        .on_action(cx.listener(|this, _: &AltHelp, _, cx| {
            if this.should_handle_option_accelerators() {
                this.apply_menu_scope(MenuCategory::Help, cx);
            }
        }))
        // Sheet navigation
        .on_action(cx.listener(|this, _: &NextSheet, _, cx| {
            this.next_sheet(cx);
        }))
        .on_action(cx.listener(|this, _: &PrevSheet, _, cx| {
            this.prev_sheet(cx);
        }))
        .on_action(cx.listener(|this, _: &AddSheet, _, cx| {
            this.add_sheet(cx);
        }))
}
