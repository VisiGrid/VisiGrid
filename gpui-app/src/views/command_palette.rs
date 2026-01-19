use std::time::Duration;
use gpui::*;
use crate::app::Spreadsheet;

/// A command that can be executed from the palette
#[derive(Clone)]
pub struct Command {
    pub name: &'static str,
    pub keywords: &'static str,  // Additional search terms
    pub shortcut: Option<&'static str>,
    pub action: fn(&mut Spreadsheet, &mut Context<Spreadsheet>),
}

/// All available commands
pub fn all_commands() -> Vec<Command> {
    vec![
        // Navigation
        Command {
            name: "Go to Cell",
            keywords: "goto jump navigate",
            shortcut: Some("Ctrl+G"),
            action: |app, cx| app.show_goto(cx),
        },
        Command {
            name: "Find in Cells",
            keywords: "search",
            shortcut: Some("Ctrl+F"),
            action: |app, cx| app.show_find(cx),
        },
        Command {
            name: "Go to Start (A1)",
            keywords: "home beginning",
            shortcut: Some("Ctrl+Home"),
            action: |app, cx| {
                app.selected = (0, 0);
                app.selection_end = None;
                app.scroll_row = 0;
                app.scroll_col = 0;
                cx.notify();
            },
        },
        Command {
            name: "Select All",
            keywords: "selection",
            shortcut: Some("Ctrl+A"),
            action: |app, cx| app.select_all(cx),
        },

        // Editing
        Command {
            name: "Fill Down",
            keywords: "copy formula",
            shortcut: Some("Ctrl+D"),
            action: |app, cx| app.fill_down(cx),
        },
        Command {
            name: "Fill Right",
            keywords: "copy formula",
            shortcut: Some("Ctrl+R"),
            action: |app, cx| app.fill_right(cx),
        },
        Command {
            name: "Clear Cells",
            keywords: "delete remove empty",
            shortcut: Some("Delete"),
            action: |app, cx| app.delete_selection(cx),
        },
        Command {
            name: "Undo",
            keywords: "revert back",
            shortcut: Some("Ctrl+Z"),
            action: |app, cx| app.undo(cx),
        },
        Command {
            name: "Redo",
            keywords: "forward",
            shortcut: Some("Ctrl+Y"),
            action: |app, cx| app.redo(cx),
        },

        // Clipboard
        Command {
            name: "Copy",
            keywords: "clipboard",
            shortcut: Some("Ctrl+C"),
            action: |app, cx| app.copy(cx),
        },
        Command {
            name: "Cut",
            keywords: "clipboard",
            shortcut: Some("Ctrl+X"),
            action: |app, cx| app.cut(cx),
        },
        Command {
            name: "Paste",
            keywords: "clipboard",
            shortcut: Some("Ctrl+V"),
            action: |app, cx| app.paste(cx),
        },

        // Formatting
        Command {
            name: "Toggle Bold",
            keywords: "format style",
            shortcut: Some("Ctrl+B"),
            action: |app, cx| app.toggle_bold(cx),
        },
        Command {
            name: "Toggle Italic",
            keywords: "format style",
            shortcut: Some("Ctrl+I"),
            action: |app, cx| app.toggle_italic(cx),
        },
        Command {
            name: "Toggle Underline",
            keywords: "format style",
            shortcut: Some("Ctrl+U"),
            action: |app, cx| app.toggle_underline(cx),
        },

        // File
        Command {
            name: "New File",
            keywords: "create workbook",
            shortcut: Some("Ctrl+N"),
            action: |app, cx| app.new_file(cx),
        },
        Command {
            name: "Open File",
            keywords: "load",
            shortcut: Some("Ctrl+O"),
            action: |app, cx| app.open_file(cx),
        },
        Command {
            name: "Save",
            keywords: "write",
            shortcut: Some("Ctrl+S"),
            action: |app, cx| app.save(cx),
        },
        Command {
            name: "Save As",
            keywords: "write export",
            shortcut: Some("Ctrl+Shift+S"),
            action: |app, cx| app.save_as(cx),
        },
        Command {
            name: "Export as CSV",
            keywords: "save",
            shortcut: None,
            action: |app, cx| app.export_csv(cx),
        },

        // Help
        Command {
            name: "Show Keyboard Shortcuts",
            keywords: "help keys bindings hotkeys",
            shortcut: None,
            action: |app, cx| {
                app.status_message = Some("Shortcuts: Ctrl+D Fill Down, Ctrl+R Fill Right, Ctrl+Enter Multi-edit".into());
                cx.notify();
            },
        },
    ]
}

/// Filter commands by query (substring match on name + keywords)
pub fn filter_commands(query: &str) -> Vec<Command> {
    let commands = all_commands();
    if query.is_empty() {
        return commands;
    }

    let query_lower = query.to_lowercase();
    commands
        .into_iter()
        .filter(|cmd| {
            cmd.name.to_lowercase().contains(&query_lower)
                || cmd.keywords.to_lowercase().contains(&query_lower)
        })
        .collect()
}

// Colors - subtle dark theme like Zed
const BG_OVERLAY: u32 = 0x00000060;        // Subtle dark overlay
const BG_PALETTE: u32 = 0x2b2d30;          // Dark gray palette background
const BG_SELECTED: u32 = 0x3c3f41;         // Subtle gray selection
const BG_HOVER: u32 = 0x35373a;            // Very subtle hover
const TEXT_PRIMARY: u32 = 0xbcbec4;        // Slightly muted white text
const TEXT_SECONDARY: u32 = 0x6f737a;      // Muted gray text
const TEXT_PLACEHOLDER: u32 = 0x5a5d63;    // Placeholder text
const BORDER_SUBTLE: u32 = 0x3c3f41;       // Subtle borders

/// Render the command palette overlay
pub fn render_command_palette(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) -> impl IntoElement {
    let filtered = filter_commands(&app.palette_query);
    let selected_idx = app.palette_selected;
    let query = app.palette_query.clone();
    let has_query = !query.is_empty();

    div()
        .absolute()
        .inset_0()
        .flex()
        .items_start()
        .justify_center()
        .pt(px(100.0))
        .bg(rgba(BG_OVERLAY))
        // Click outside to close
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, _, cx| {
            this.hide_palette(cx);
        }))
        .child(
            div()
                .w(px(500.0))
                .max_h(px(380.0))
                .bg(rgb(BG_PALETTE))
                .rounded_md()
                .shadow_lg()
                .overflow_hidden()
                .flex()
                .flex_col()
                // Stop click propagation on the palette itself
                .on_mouse_down(MouseButton::Left, |_, _, cx| {
                    cx.stop_propagation();
                })
                // Search input
                .child(
                    div()
                        .flex()
                        .items_center()
                        .px_3()
                        .py(px(10.0))
                        .border_b_1()
                        .border_color(rgb(BORDER_SUBTLE))
                        // Input area with cursor
                        .child(
                            div()
                                .flex_1()
                                .flex()
                                .items_center()
                                .child(
                                    div()
                                        .text_color(if has_query { rgb(TEXT_PRIMARY) } else { rgb(TEXT_PLACEHOLDER) })
                                        .text_size(px(13.0))
                                        .child(if has_query {
                                            query.clone()
                                        } else {
                                            "Execute a command...".to_string()
                                        })
                                )
                                // Blinking cursor
                                .child(
                                    div()
                                        .w(px(1.0))
                                        .h(px(14.0))
                                        .bg(rgb(TEXT_PRIMARY))
                                        .ml(px(1.0))
                                        .with_animation(
                                            "cursor-blink",
                                            Animation::new(Duration::from_millis(530))
                                                .repeat()
                                                .with_easing(pulsating_between(0.0, 1.0)),
                                            |div, delta| {
                                                let opacity = if delta > 0.5 { 0.0 } else { 1.0 };
                                                div.opacity(opacity)
                                            },
                                        )
                                )
                        )
                )
                // Command list
                .child({
                    let list = div()
                        .flex_1()
                        .overflow_hidden()
                        .py_1()
                        .children(
                            filtered.iter().enumerate().take(12).map(|(idx, cmd)| {
                                let is_selected = idx == selected_idx;
                                render_command_item(cmd, is_selected, idx, cx)
                            })
                        );
                    if filtered.is_empty() {
                        list.child(
                            div()
                                .px_4()
                                .py_6()
                                .text_color(rgb(TEXT_SECONDARY))
                                .text_size(px(14.0))
                                .child("No matching commands")
                        )
                    } else {
                        list
                    }
                })
                // Footer with hints
                .child(
                    div()
                        .flex()
                        .items_center()
                        .justify_end()
                        .gap_3()
                        .px_3()
                        .py(px(6.0))
                        .border_t_1()
                        .border_color(rgb(BORDER_SUBTLE))
                        .text_size(px(11.0))
                        .text_color(rgb(TEXT_SECONDARY))
                        .child("Run")
                        .child(
                            div()
                                .text_color(rgb(TEXT_PLACEHOLDER))
                                .child("enter")
                        )
                )
        )
}

fn render_command_item(
    cmd: &Command,
    is_selected: bool,
    idx: usize,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    let action = cmd.action;
    let name = cmd.name;
    let shortcut = cmd.shortcut;

    let bg_color = if is_selected { rgb(BG_SELECTED) } else { rgba(0x00000000) };

    let mut item = div()
        .id(ElementId::NamedInteger("palette-cmd".into(), idx as u64))
        .flex()
        .items_center()
        .justify_between()
        .px_3()
        .py(px(6.0))
        .cursor_pointer()
        .bg(bg_color)
        .on_mouse_down(MouseButton::Left, cx.listener(move |this, _, _, cx| {
            action(this, cx);
            this.hide_palette(cx);
        }))
        .child(
            div()
                .text_color(rgb(TEXT_PRIMARY))
                .text_size(px(13.0))
                .child(name)
        );

    if !is_selected {
        item = item.hover(|s| s.bg(rgb(BG_HOVER)));
    }

    if let Some(shortcut_text) = shortcut {
        item = item.child(
            div()
                .text_color(rgb(TEXT_SECONDARY))
                .text_size(px(12.0))
                .child(shortcut_text)
        );
    }

    item
}
