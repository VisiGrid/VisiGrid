//! Lua console panel view.
//!
//! A REPL panel for executing Lua scripts. Shows at the bottom of the window
//! when toggled with Ctrl+Shift+L.
//!
//! ## Virtual Scroll
//!
//! The console uses virtual scrolling to handle large outputs efficiently.
//! PageUp/PageDown navigate through output history. Ctrl+Home/End jump to
//! start/end. The view auto-scrolls to bottom when new output arrives.

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::app::Spreadsheet;
use crate::scripting::{OutputEntry, OutputKind, SheetSnapshot, LuaOp, LuaCellValue, MIN_CONSOLE_HEIGHT, MAX_CONSOLE_HEIGHT};
use crate::scripting::examples::{EXAMPLES, get_example, find_example};
use crate::theme::TokenKey;

/// Render the Lua console panel (if visible)
pub fn render_lua_console(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let console = &app.lua_console;

    if !console.visible {
        return div().into_any_element();
    }

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let error_color = app.token(TokenKey::Error);
    let editor_bg = app.token(TokenKey::EditorBg);

    let console_height = console.height;

    div()
        .id("lua-console-panel")
        .key_context("LuaConsole")
        .track_focus(&app.console_focus_handle)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            window.focus(&this.console_focus_handle, cx);
            cx.notify();
        }))
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            handle_console_key_from_main(this, event, window, cx);
            cx.stop_propagation();
        }))
        .flex_shrink_0()
        .h(px(console_height))
        .bg(panel_bg)
        .border_t_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .child(
            // Resize handle at top
            div()
                .id("lua-console-resize")
                .h(px(4.0))
                .w_full()
                .cursor(CursorStyle::ResizeUpDown)
                .on_mouse_down(MouseButton::Left, cx.listener(|this, event: &MouseDownEvent, _, cx| {
                    this.lua_console.resizing = true;
                    this.lua_console.resize_start_y = event.position.y.into();
                    this.lua_console.resize_start_height = this.lua_console.height;
                    cx.notify();
                }))
        )
        .child(
            // Header bar
            div()
                .h(px(24.0))
                .px_2()
                .flex()
                .items_center()
                .justify_between()
                .border_b_1()
                .border_color(panel_border)
                .child(
                    div()
                        .text_xs()
                        .text_color(text_muted)
                        .child("Lua Console")
                )
                .child(
                    // Close button
                    div()
                        .id("lua-console-close")
                        .px_1()
                        .cursor_pointer()
                        .text_xs()
                        .text_color(text_muted)
                        .hover(|s| s.text_color(text_primary))
                        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
                            this.lua_console.hide();
                            cx.notify();
                        }))
                        .child("×")
                )
        )
        .child(
            // Output area with virtual scroll
            div()
                .flex_1()
                .flex()
                .flex_col()
                .overflow_hidden()
                .child(
                    // Output lines
                    div()
                        .flex_1()
                        .overflow_hidden()
                        .px_2()
                        .py_1()
                        .when(console.output.is_empty(), |d| {
                            d.child(
                                div()
                                    .text_xs()
                                    .text_color(text_muted.opacity(0.5))
                                    .child("Output appears here. Press Enter to run.")
                            )
                        })
                        .children(
                            console.visible_output().iter().enumerate().map(|(i, entry)| {
                                render_output_entry(entry, console.view_start + i, text_primary, text_muted, accent, error_color)
                            })
                        )
                )
                .when(console.scroll_info().is_some(), |d| {
                    d.child(
                        // Scroll indicator bar
                        div()
                            .h(px(16.0))
                            .px_2()
                            .flex()
                            .items_center()
                            .justify_between()
                            .border_t_1()
                            .border_color(panel_border)
                            .bg(editor_bg)
                            .child(
                                // Scroll position info
                                div()
                                    .text_xs()
                                    .text_color(text_muted)
                                    .child(console.scroll_info().unwrap_or_default())
                            )
                            .child(
                                // Scroll controls
                                div()
                                    .flex()
                                    .gap_1()
                                    .child(
                                        scroll_button("▲", console.can_scroll_up(), text_muted, text_primary, cx, |this, cx| {
                                            this.lua_console.scroll_page_up();
                                            cx.notify();
                                        })
                                    )
                                    .child(
                                        scroll_button("▼", console.can_scroll_down(), text_muted, text_primary, cx, |this, cx| {
                                            this.lua_console.scroll_page_down();
                                            cx.notify();
                                        })
                                    )
                            )
                    )
                })
        )
        .child(
            // Input area
            div()
                .h(px(28.0))
                .px_2()
                .flex()
                .items_center()
                .gap_2()
                .border_t_1()
                .border_color(panel_border)
                .child(
                    div()
                        .text_xs()
                        .text_color(accent)
                        .child(">")
                )
                .child(
                    // Input display (we'll render the input text here)
                    render_input_area(app, editor_bg, text_primary, accent, cx)
                )
        )
        .child(
            // Hint footer
            div()
                .h(px(16.0))
                .px_2()
                .flex()
                .items_center()
                .justify_end()
                .text_xs()
                .text_color(text_muted.opacity(0.6))
                .child("Ctrl+Enter to run · Undo reverses the entire script")
        )
        .into_any_element()
}

/// Render a single output entry
fn render_output_entry(
    entry: &OutputEntry,
    index: usize,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    error_color: Hsla,
) -> Stateful<Div> {
    let (prefix, color) = match entry.kind {
        OutputKind::Input => ("> ", text_muted),
        OutputKind::Result => ("", accent),
        OutputKind::Print => ("", text_primary),
        OutputKind::Error => ("", error_color),
        OutputKind::System => ("", text_muted),
    };

    div()
        .id(ElementId::Name(format!("lua-output-{}", index).into()))
        .text_xs()
        .font_family("monospace")
        .text_color(color)
        .child(format!("{}{}", prefix, entry.text))
}

/// Render a scroll button (▲ or ▼)
fn scroll_button<F>(
    label: &'static str,
    enabled: bool,
    text_muted: Hsla,
    text_primary: Hsla,
    cx: &mut Context<Spreadsheet>,
    on_click: F,
) -> Stateful<Div>
where
    F: Fn(&mut Spreadsheet, &mut Context<Spreadsheet>) + 'static,
{
    let color = if enabled { text_muted } else { text_muted.opacity(0.3) };

    div()
        .id(ElementId::Name(format!("scroll-btn-{}", label).into()))
        .px_1()
        .text_xs()
        .text_color(color)
        .when(enabled, |div| {
            div
                .cursor_pointer()
                .hover(|s| s.text_color(text_primary))
                .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
                    on_click(this, cx);
                }))
        })
        .child(label)
}

/// Render the input area with cursor
fn render_input_area(
    app: &Spreadsheet,
    editor_bg: Hsla,
    text_primary: Hsla,
    accent: Hsla,
    _cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let console = &app.lua_console;
    let input = &console.input;
    let cursor = console.cursor;

    // Split input at cursor for rendering
    let before = &input[..cursor];
    let after = &input[cursor..];

    div()
        .id("lua-input")
        .flex_1()
        .h_full()
        .px_1()
        .bg(editor_bg)
        .rounded_sm()
        .flex()
        .items_center()
        .text_xs()
        .font_family("monospace")
        .text_color(text_primary)
        .child(before.to_string())
        .child(
            // Cursor
            div()
                .w(px(1.0))
                .h(px(12.0))
                .bg(accent)
        )
        .child(after.to_string())
}

/// Handle keyboard input in the console (called from main key handler)
pub fn handle_console_key_from_main(app: &mut Spreadsheet, event: &KeyDownEvent, window: &mut Window, cx: &mut Context<Spreadsheet>) {
    let key = &event.keystroke.key;
    let ctrl = event.keystroke.modifiers.control;
    let shift = event.keystroke.modifiers.shift;

    match key.as_str() {
        "enter" => {
            if ctrl {
                // Ctrl+Enter: always execute (power user shortcut)
                execute_console(app, cx);
            } else if shift {
                // Shift+Enter: insert newline for multiline input
                app.lua_console.insert("\n");
            } else {
                // Enter: execute
                execute_console(app, cx);
            }
        }
        "escape" => {
            // Cancel / close and restore focus to main grid
            app.lua_console.hide();
            window.focus(&app.focus_handle, cx);
        }
        "backspace" => {
            app.lua_console.backspace();
        }
        "delete" => {
            app.lua_console.delete();
        }
        "left" => {
            app.lua_console.cursor_left();
        }
        "right" => {
            app.lua_console.cursor_right();
        }
        "home" => {
            if ctrl {
                // Ctrl+Home: scroll to top of output
                app.lua_console.scroll_to_start();
            } else {
                app.lua_console.cursor_home();
            }
        }
        "end" => {
            if ctrl {
                // Ctrl+End: scroll to bottom of output
                app.lua_console.scroll_to_end();
            } else {
                app.lua_console.cursor_end();
            }
        }
        "up" => {
            app.lua_console.history_prev();
        }
        "down" => {
            app.lua_console.history_next();
        }
        "pageup" => {
            // PageUp: scroll output up
            app.lua_console.scroll_page_up();
        }
        "pagedown" => {
            // PageDown: scroll output down
            app.lua_console.scroll_page_down();
        }
        "l" if ctrl => {
            // Ctrl+L: clear output
            app.lua_console.clear_output();
        }
        _ => {
            // Insert character
            if let Some(key_char) = &event.keystroke.key_char {
                if !ctrl && !event.keystroke.modifiers.alt {
                    app.lua_console.insert(key_char);
                }
            }
        }
    }

    cx.notify();
}

/// Execute the current input (public for action handlers)
pub fn execute_console(app: &mut Spreadsheet, cx: &mut Context<Spreadsheet>) {
    use std::time::Instant;

    let input = app.lua_console.consume_input();
    if input.trim().is_empty() {
        return;
    }

    // Echo input
    app.lua_console.push_output(OutputEntry::input(&input));

    // Handle special commands
    let trimmed = input.trim();

    if trimmed == "clear" {
        app.lua_console.clear_output();
        cx.notify();
        return;
    }

    if trimmed == "help" {
        app.lua_console.push_output(OutputEntry::system("Commands:"));
        app.lua_console.push_output(OutputEntry::system("  clear      - Clear output"));
        app.lua_console.push_output(OutputEntry::system("  examples   - List example scripts"));
        app.lua_console.push_output(OutputEntry::system("  example N  - Run example N (by number or name)"));
        app.lua_console.push_output(OutputEntry::system("  show N     - Show code for example N"));
        app.lua_console.push_output(OutputEntry::system("  help       - Show this help"));
        app.lua_console.push_output(OutputEntry::system(""));
        app.lua_console.push_output(OutputEntry::system("Keyboard:"));
        app.lua_console.push_output(OutputEntry::system("  Enter        - Execute"));
        app.lua_console.push_output(OutputEntry::system("  Shift+Enter  - Newline (multiline input)"));
        app.lua_console.push_output(OutputEntry::system("  Up/Down      - History navigation"));
        app.lua_console.push_output(OutputEntry::system("  PageUp/Down  - Scroll output"));
        app.lua_console.push_output(OutputEntry::system("  Ctrl+L       - Clear output"));
        app.lua_console.push_output(OutputEntry::system("  Escape       - Close console"));
        cx.notify();
        return;
    }

    if trimmed == "examples" {
        app.lua_console.push_output(OutputEntry::system("Example scripts (use 'example N' to load):"));
        for (i, example) in EXAMPLES.iter().enumerate() {
            app.lua_console.push_output(OutputEntry::system(
                format!("  {}. {} - {}", i + 1, example.name, example.description)
            ));
        }
        cx.notify();
        return;
    }

    // show N - preview example code without running
    if trimmed.starts_with("show ") {
        let arg = trimmed.strip_prefix("show ").unwrap().trim();

        let example = if let Ok(n) = arg.parse::<usize>() {
            get_example(n.saturating_sub(1))
        } else {
            find_example(arg)
        };

        if let Some(ex) = example {
            app.lua_console.push_output(OutputEntry::system(
                format!("-- {} --", ex.name)
            ));
            // Show code line by line
            for line in ex.code.lines() {
                app.lua_console.push_output(OutputEntry::print(line.to_string()));
            }
            app.lua_console.push_output(OutputEntry::system(
                format!("Type 'example {}' to run", arg)
            ));
        } else {
            app.lua_console.push_output(OutputEntry::error(
                format!("Unknown example: '{}'. Type 'examples' to see list.", arg)
            ));
        }
        cx.notify();
        return;
    }

    // example N - run example directly
    let code_to_run = if trimmed.starts_with("example ") {
        let arg = trimmed.strip_prefix("example ").unwrap().trim();

        let example = if let Ok(n) = arg.parse::<usize>() {
            get_example(n.saturating_sub(1))
        } else {
            find_example(arg)
        };

        if let Some(ex) = example {
            // Show which example is running
            app.lua_console.push_output(OutputEntry::system(
                format!("Running '{}'...", ex.name)
            ));
            Some(ex.code.to_string())
        } else {
            app.lua_console.push_output(OutputEntry::error(
                format!("Unknown example: '{}'. Type 'examples' to see list.", arg)
            ));
            cx.notify();
            return;
        }
    } else {
        None
    };

    // Use example code if running an example, otherwise use the user input
    let code = code_to_run.unwrap_or_else(|| input.clone());

    // Create snapshot of current sheet for Lua to read from
    let snapshot = SheetSnapshot::from_sheet(app.sheet(cx));
    let sheet_index = app.sheet_index(cx);

    // Compute selection bounds (normalize to start <= end)
    let (anchor_row, anchor_col) = app.view_state.selected;
    let (end_row, end_col) = app.view_state.selection_end.unwrap_or(app.view_state.selected);
    let selection = (
        anchor_row.min(end_row),
        anchor_col.min(end_col),
        anchor_row.max(end_row),
        anchor_col.max(end_col),
    );

    // Time the execution
    let start = Instant::now();

    // Evaluate with sheet access and selection
    let result = app.lua_runtime.eval_with_sheet_and_selection(&code, Box::new(snapshot), selection);

    let elapsed = start.elapsed();

    // Add output
    for line in &result.output {
        app.lua_console.push_output(OutputEntry::print(line.clone()));
    }

    if let Some(ref returned) = result.returned {
        app.lua_console.push_output(OutputEntry::result(returned.clone()));
    }

    if let Some(ref error) = result.error {
        app.lua_console.push_output(OutputEntry::error(error.clone()));
    }

    // Apply operations if any (single undo entry for all)
    if result.has_mutations() {
        let changes = apply_lua_ops(app, sheet_index, &result.ops, cx);
        if !changes.is_empty() {
            app.history.record_batch(sheet_index, changes);
            app.is_modified = true;
        }
    }

    // Show execution stats: ops + cells + time
    let stats = format!(
        "ops: {} | cells: {} | time: {:.1}ms",
        result.ops.len(),
        result.mutations,
        elapsed.as_secs_f64() * 1000.0
    );
    app.lua_console.push_output(OutputEntry::system(stats));

    cx.notify();
}

/// Apply Lua operations to the sheet and return undo changes.
/// Uses batched, tracked mutations so dependents recalculate once at the end.
fn apply_lua_ops(
    app: &mut Spreadsheet,
    sheet_index: usize,
    ops: &[LuaOp],
    cx: &mut gpui::Context<Spreadsheet>,
) -> Vec<crate::history::CellChange> {
    use crate::history::CellChange;

    if ops.is_empty() {
        return Vec::new();
    }

    app.workbook.update(cx, |wb, _| {
        let mut guard = wb.batch_guard();
        let mut changes = Vec::new();

        for op in ops {
            match op {
                LuaOp::SetValue { row, col, value } => {
                    let row = *row as usize;
                    let col = *col as usize;

                    let old_value = guard.sheet(sheet_index)
                        .map(|s| s.get_raw(row, col))
                        .unwrap_or_default();

                    let new_value = lua_cell_value_to_string(value);
                    guard.set_cell_value_tracked(sheet_index, row, col, &new_value);

                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value,
                    });
                }
                LuaOp::SetFormula { row, col, formula } => {
                    let row = *row as usize;
                    let col = *col as usize;

                    let old_value = guard.sheet(sheet_index)
                        .map(|s| s.get_raw(row, col))
                        .unwrap_or_default();

                    guard.set_cell_value_tracked(sheet_index, row, col, formula);

                    changes.push(CellChange {
                        row,
                        col,
                        old_value,
                        new_value: formula.clone(),
                    });
                }
            }
        }

        changes
    })
}

/// Convert LuaCellValue to a string suitable for sheet.set_value()
fn lua_cell_value_to_string(value: &LuaCellValue) -> String {
    match value {
        LuaCellValue::Nil => String::new(),  // Empty clears the cell
        LuaCellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{:.0}", n)
            } else {
                format!("{}", n)
            }
        }
        LuaCellValue::String(s) => s.clone(),
        LuaCellValue::Bool(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        LuaCellValue::Error(e) => format!("#ERROR: {}", e),
    }
}
