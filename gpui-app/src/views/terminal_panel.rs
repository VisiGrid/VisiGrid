//! Terminal panel view.
//!
//! Renders a PTY terminal at the bottom of the window. Uses `alacritty_terminal`
//! for ANSI grid state and renders line-by-line with styled runs.

use std::path::{Path, PathBuf};

use gpui::*;
use gpui::prelude::FluentBuilder;

use alacritty_terminal::grid::{Dimensions, Scroll};
use alacritty_terminal::index::{Column, Line};
use alacritty_terminal::term::cell::Flags as CellFlags;
use alacritty_terminal::vte::ansi::{Color as AnsiColor, NamedColor};

use crate::app::Spreadsheet;
use crate::terminal::state::MAX_TERMINAL_HEIGHT;
use crate::theme::TokenKey;

/// Format a path for display in the breadcrumb.
/// Replaces home dir with `~`, clamps to last 3 segments with ellipsis.
fn format_cwd_display(path: &std::path::Path) -> String {
    let mut display = path.to_string_lossy().to_string();

    // Replace home directory with ~
    if let Some(home) = dirs::home_dir() {
        let home_str = home.to_string_lossy().to_string();
        if display.starts_with(&home_str) {
            display = format!("~{}", &display[home_str.len()..]);
        }
    }

    // Clamp to last 3 path segments with ellipsis
    let parts: Vec<&str> = display.split('/').filter(|s| !s.is_empty()).collect();
    if parts.len() > 3 {
        format!(".../{}", parts[parts.len()-3..].join("/"))
    } else {
        display
    }
}

/// Font metrics for the terminal grid.
const TERM_FONT_SIZE: f32 = 15.0;
const TERM_CELL_WIDTH: f32 = 9.0;   // Monospace approximate at 15px
const TERM_CELL_HEIGHT: f32 = 20.0;  // Line height
pub const TERM_FONT_FAMILY: &str = "CaskaydiaMono Nerd Font";

/// Render the terminal panel (if visible).
pub fn render_terminal_panel(
    app: &Spreadsheet,
    _window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) -> impl IntoElement {
    if !app.terminal.visible {
        return div().into_any_element();
    }

    let panel_bg = app.token(TokenKey::PanelBg);
    let panel_border = app.token(TokenKey::PanelBorder);
    let text_primary = app.token(TokenKey::TextPrimary);
    let text_muted = app.token(TokenKey::TextMuted);
    let accent = app.token(TokenKey::Accent);
    let editor_bg = app.token(TokenKey::EditorBg);

    // Viewport-relative max height: cap at 60% of window height
    let window_height: f32 = app.window_size.height.into();
    let effective_max = if window_height > 0.0 {
        MAX_TERMINAL_HEIGHT.min(window_height * 0.6)
    } else {
        MAX_TERMINAL_HEIGHT
    };

    let terminal_height = if app.terminal.is_maximized {
        effective_max
    } else {
        app.terminal.height
    };

    let is_maximized = app.terminal.is_maximized;
    let exited = app.terminal.exited;
    let exit_code = app.terminal.exit_code;

    // Build the terminal grid content
    let (grid_content, display_offset) = if let Some(ref term_arc) = app.terminal.term {
        // Lock the term, extract renderable content, then drop the lock before building elements
        let term = term_arc.lock();
        let content = term.renderable_content();
        let cols = term.columns();
        let lines = term.screen_lines();
        let display_offset = content.display_offset;
        let cursor = content.cursor;
        let mode = content.mode;
        let show_cursor = mode.contains(alacritty_terminal::term::TermMode::SHOW_CURSOR)
            && display_offset == 0; // Hide cursor when scrolled up

        // Collect row data while holding the lock
        let mut row_data: Vec<Vec<(char, AnsiColor, AnsiColor, CellFlags)>> = Vec::with_capacity(lines);
        for line_idx in 0..lines {
            let line = Line(line_idx as i32);
            let row = &term.grid()[line];
            let mut cells = Vec::with_capacity(cols);
            for col_idx in 0..cols {
                let cell = &row[Column(col_idx)];
                cells.push((cell.c, cell.fg, cell.bg, cell.flags));
            }
            row_data.push(cells);
        }

        let cursor_line = cursor.point.line.0 as usize;
        let cursor_col = cursor.point.column.0;

        drop(term); // Release lock before building element tree

        // Build element tree from extracted data
        let mut rows_el = div()
            .flex()
            .flex_col()
            .w_full()
            .overflow_hidden();

        for (line_idx, cells) in row_data.iter().enumerate() {
            let mut spans: Vec<AnyElement> = Vec::new();
            let mut run_start = 0;
            let is_cursor_line = show_cursor && line_idx == cursor_line;

            while run_start < cells.len() {
                let (_, fg, bg, flags) = cells[run_start];

                // Find end of run with same style
                let mut run_end = run_start + 1;
                while run_end < cells.len() {
                    let (_, nfg, nbg, nflags) = cells[run_end];
                    if nfg != fg || nbg != bg || nflags != flags {
                        break;
                    }
                    run_end += 1;
                }

                // If the cursor falls inside this run, split the run at cursor_col
                // so the cursor element is rendered inline at the correct position.
                if is_cursor_line && cursor_col >= run_start && cursor_col < run_end {
                    // Part before cursor
                    if cursor_col > run_start {
                        let text: String = cells[run_start..cursor_col]
                            .iter()
                            .map(|(c, _, _, f)| cell_char(*c, *f))
                            .collect();
                        spans.push(make_span(text, fg, bg, flags, text_primary, editor_bg).into_any_element());
                    }

                    // Cursor cell — rendered inline with inverted colors
                    let cursor_char = cell_char(cells[cursor_col].0, cells[cursor_col].3);
                    spans.push(
                        div()
                            .font_family(TERM_FONT_FAMILY)
                            .bg(text_primary.opacity(0.7))
                            .text_color(editor_bg)
                            .text_size(px(TERM_FONT_SIZE))
                            .child(cursor_char.to_string())
                            .into_any_element()
                    );

                    // Part after cursor
                    if cursor_col + 1 < run_end {
                        let text: String = cells[cursor_col + 1..run_end]
                            .iter()
                            .map(|(c, _, _, f)| cell_char(*c, *f))
                            .collect();
                        spans.push(make_span(text, fg, bg, flags, text_primary, editor_bg).into_any_element());
                    }
                } else {
                    // Normal run — no cursor intersection
                    let text: String = cells[run_start..run_end]
                        .iter()
                        .map(|(c, _, _, f)| cell_char(*c, *f))
                        .collect();
                    spans.push(make_span(text, fg, bg, flags, text_primary, editor_bg).into_any_element());
                }

                run_start = run_end;
            }

            // If cursor is past all cells on this line, append cursor at end
            if is_cursor_line && cursor_col >= cells.len() {
                spans.push(
                    div()
                        .font_family(TERM_FONT_FAMILY)
                        .bg(text_primary.opacity(0.7))
                        .text_color(editor_bg)
                        .text_size(px(TERM_FONT_SIZE))
                        .child(" ")
                        .into_any_element()
                );
            }

            let row_el = div()
                .flex()
                .flex_row()
                .h(px(TERM_CELL_HEIGHT))
                .w_full()
                .children(spans);

            rows_el = rows_el.child(row_el);
        }

        (div()
            .flex()
            .flex_col()
            .size_full()
            .relative()
            .child(rows_el)
            .into_any_element(),
        display_offset)
    } else if exited {
        let msg = match exit_code {
            Some(code) => format!("Process exited with code {}. Press Enter to restart.", code),
            None => "Process exited. Press Enter to restart.".to_string(),
        };
        (div()
            .flex()
            .items_center()
            .justify_center()
            .size_full()
            .text_color(text_muted)
            .text_size(px(TERM_FONT_SIZE))
            .child(msg)
            .into_any_element(),
        0)
    } else {
        (div()
            .flex()
            .items_center()
            .justify_center()
            .size_full()
            .text_color(text_muted)
            .text_size(px(TERM_FONT_SIZE))
            .child("Starting terminal...")
            .into_any_element(),
        0)
    };

    let is_scrolled = display_offset > 0;

    // CWD breadcrumb (for copy-on-click)
    let cwd_display = app.terminal.workspace_root.as_ref().map(|p| format_cwd_display(p));
    let cwd_full = app.terminal.workspace_root.clone();

    // Header bar
    let header = div()
        .flex()
        .flex_row()
        .items_center()
        .justify_between()
        .h(px(28.0))
        .px(px(8.0))
        .border_b_1()
        .border_color(panel_border)
        .child(
            div()
                .flex()
                .flex_row()
                .items_center()
                .gap(px(8.0))
                .child(
                    div()
                        .text_color(accent)
                        .text_size(px(11.0))
                        .font_weight(FontWeight::SEMIBOLD)
                        .child("Terminal")
                )
                .when(exited, |d| {
                    d.child(
                        div()
                            .text_color(text_muted)
                            .text_size(px(11.0))
                            .child("(exited)")
                    )
                })
                .when(is_scrolled, |d| {
                    d.child(
                        div()
                            .id("terminal-scroll-indicator")
                            .cursor_pointer()
                            .text_color(text_muted)
                            .text_size(px(11.0))
                            .hover(|s| s.text_color(text_primary))
                            .child(format!("(Scrolled +{}) · Press End to follow", display_offset))
                            .on_click(cx.listener(|this, _, _, cx| {
                                // Click to jump to bottom
                                if let Some(ref term_arc) = this.terminal.term {
                                    let mut term = term_arc.lock();
                                    term.grid_mut().scroll_display(Scroll::Bottom);
                                }
                                cx.notify();
                            }))
                    )
                })
        )
        .child(
            div()
                .flex()
                .flex_row()
                .items_center()
                .gap(px(8.0))
                // CWD breadcrumb (click to copy full path)
                .when_some(cwd_display, |d, cwd_str| {
                    d.child(
                        div()
                            .id("terminal-cwd")
                            .cursor_pointer()
                            .text_color(text_muted)
                            .text_size(px(10.0))
                            .hover(|s| s.text_color(text_primary))
                            .child(cwd_str)
                            .on_click(cx.listener(move |_this, _, _, cx| {
                                if let Some(ref path) = cwd_full {
                                    let full = path.display().to_string();
                                    cx.write_to_clipboard(ClipboardItem::new_string(full));
                                }
                            }))
                    )
                })
                // Maximize button
                .child(
                    div()
                        .id("terminal-maximize")
                        .cursor_pointer()
                        .text_color(text_muted)
                        .text_size(px(11.0))
                        .px(px(4.0))
                        .hover(|s| s.text_color(text_primary))
                        .child(if is_maximized { "Restore" } else { "Maximize" })
                        .on_click(cx.listener(move |this, _, _, cx| {
                            let window_height: f32 = this.window_size.height.into();
                            let effective_max = if window_height > 0.0 {
                                MAX_TERMINAL_HEIGHT.min(window_height * 0.6)
                            } else {
                                MAX_TERMINAL_HEIGHT
                            };
                            this.terminal.toggle_maximize(effective_max);
                            cx.notify();
                        }))
                )
                // Close button
                .child(
                    div()
                        .id("terminal-close")
                        .cursor_pointer()
                        .text_color(text_muted)
                        .text_size(px(14.0))
                        .px(px(4.0))
                        .hover(|s| s.text_color(text_primary))
                        .child("×")
                        .on_click(cx.listener(|this, _, _, cx| {
                            this.terminal.hide();
                            cx.notify();
                        }))
                )
        );

    // Resize handle (3px strip at top of panel)
    let resize_handle = div()
        .id("terminal-resize-handle")
        .h(px(3.0))
        .w_full()
        .cursor(CursorStyle::ResizeUpDown)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, event: &MouseDownEvent, _, cx| {
            this.terminal.resizing = true;
            this.terminal.resize_start_y = event.position.y.into();
            this.terminal.resize_start_height = this.terminal.height;
            cx.notify();
        }));

    div()
        .id("terminal-panel")
        .key_context("Terminal")
        .track_focus(&app.terminal_focus_handle)
        .on_mouse_down(MouseButton::Left, cx.listener(|this, _, window, cx| {
            window.focus(&this.terminal_focus_handle, cx);
            cx.notify();
        }))
        .on_key_down(cx.listener(|this, event: &KeyDownEvent, window, cx| {
            handle_terminal_key(this, event, window, cx);
            cx.stop_propagation();
        }))
        .on_scroll_wheel(cx.listener(|this, event: &ScrollWheelEvent, _, cx| {
            handle_terminal_scroll(this, event, cx);
        }))
        .on_drop(cx.listener(|this, paths: &ExternalPaths, window, cx| {
            handle_terminal_drop(this, paths.paths(), window, cx);
        }))
        .flex_shrink_0()
        .h(px(terminal_height))
        .bg(panel_bg)
        .border_t_1()
        .border_color(panel_border)
        .flex()
        .flex_col()
        .child(resize_handle)
        .child(header)
        .child(
            div()
                .flex_1()
                .overflow_hidden()
                .bg(editor_bg)
                .p(px(4.0))
                .font_family(TERM_FONT_FAMILY)
                .child(grid_content)
        )
        .into_any_element()
}

/// Handle keyboard input in the terminal panel.
fn handle_terminal_key(
    app: &mut Spreadsheet,
    event: &KeyDownEvent,
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) {
    // If terminal has exited, Enter restarts it
    if app.terminal.exited {
        if event.keystroke.key == "enter" {
            app.terminal.exited = false;
            app.terminal.exit_code = None;
            app.spawn_terminal(window, cx);
            cx.notify();
        }
        return;
    }

    if app.terminal.term.is_none() {
        return;
    }

    let mods = &event.keystroke.modifiers;
    let key = event.keystroke.key.as_str();

    // Ctrl+Shift+C: copy viewport text (no selection yet)
    // Ctrl+Shift+V: paste from clipboard
    if mods.control && mods.shift {
        match key {
            "c" | "C" => {
                copy_viewport_text(app, cx);
                return;
            }
            "v" | "V" => {
                if let Some(item) = cx.read_from_clipboard() {
                    if let Some(text) = item.text() {
                        app.terminal.write_to_pty(text.as_bytes());
                        cx.notify();
                    }
                }
                return;
            }
            _ => {}
        }
    }

    // Control characters (Ctrl+letter) — Ctrl+C sends SIGINT here
    if mods.control && !mods.shift && !mods.alt {
        if key.len() == 1 {
            let ch = key.chars().next().unwrap();
            if ch.is_ascii_alphabetic() {
                let ctrl_byte = (ch.to_ascii_uppercase() as u8) - b'A' + 1;
                app.terminal.write_to_pty(&[ctrl_byte]);
                cx.notify();
                return;
            }
        }
    }

    // Special keys
    let bytes: Option<&[u8]> = match key {
        "enter" => Some(b"\r"),
        "backspace" => Some(b"\x7f"),
        "tab" => Some(b"\t"),
        "escape" => Some(b"\x1b"),
        "up" => Some(b"\x1b[A"),
        "down" => Some(b"\x1b[B"),
        "right" => Some(b"\x1b[C"),
        "left" => Some(b"\x1b[D"),
        "home" => Some(b"\x1b[H"),
        "end" => {
            // End key: if scrolled up, jump to bottom instead of sending to PTY
            if let Some(ref term_arc) = app.terminal.term {
                let term = term_arc.lock();
                if term.grid().display_offset() > 0 {
                    drop(term);
                    let mut term = app.terminal.term.as_ref().unwrap().lock();
                    term.grid_mut().scroll_display(Scroll::Bottom);
                    cx.notify();
                    return;
                }
            }
            Some(b"\x1b[F")
        }
        "delete" => Some(b"\x1b[3~"),
        "pageup" => Some(b"\x1b[5~"),
        "pagedown" => Some(b"\x1b[6~"),
        "f1" => Some(b"\x1bOP"),
        "f2" => Some(b"\x1bOQ"),
        "f3" => Some(b"\x1bOR"),
        "f4" => Some(b"\x1bOS"),
        "f5" => Some(b"\x1b[15~"),
        "f6" => Some(b"\x1b[17~"),
        "f7" => Some(b"\x1b[18~"),
        "f8" => Some(b"\x1b[19~"),
        "f9" => Some(b"\x1b[20~"),
        "f10" => Some(b"\x1b[21~"),
        "f11" => Some(b"\x1b[23~"),
        "f12" => Some(b"\x1b[24~"),
        _ => None,
    };

    if let Some(bytes) = bytes {
        app.terminal.write_to_pty(bytes);
        cx.notify();
        return;
    }

    // Regular character input
    if let Some(ref key_char) = event.keystroke.key_char {
        if !mods.control && !mods.alt {
            app.terminal.write_to_pty(key_char.as_bytes());
            cx.notify();
        }
    }
}

/// Handle mouse wheel scroll in the terminal panel.
fn handle_terminal_scroll(
    app: &mut Spreadsheet,
    event: &ScrollWheelEvent,
    cx: &mut Context<Spreadsheet>,
) {
    if let Some(ref term_arc) = app.terminal.term {
        let delta = event.delta.pixel_delta(px(TERM_CELL_HEIGHT));
        let dy: f32 = delta.y.into();
        let lines = (dy / TERM_CELL_HEIGHT).round() as i32;

        if lines != 0 {
            let mut term = term_arc.lock();
            term.grid_mut().scroll_display(Scroll::Delta(lines));
            drop(term);
            cx.notify();
        }
    }
}

/// Copy the visible viewport text to clipboard.
///
/// Extracts all visible rows as plain text, trimming trailing whitespace per line.
fn copy_viewport_text(app: &Spreadsheet, cx: &mut Context<Spreadsheet>) {
    let Some(ref term_arc) = app.terminal.term else { return };

    let term = term_arc.lock();
    let cols = term.columns();
    let lines = term.screen_lines();

    let mut text = String::new();
    for line_idx in 0..lines {
        let line = Line(line_idx as i32);
        let row = &term.grid()[line];
        let mut row_text = String::with_capacity(cols);
        for col_idx in 0..cols {
            let cell = &row[Column(col_idx)];
            if cell.flags.contains(CellFlags::WIDE_CHAR_SPACER)
                || cell.flags.contains(CellFlags::LEADING_WIDE_CHAR_SPACER)
            {
                continue; // Skip spacer cells
            }
            row_text.push(cell.c);
        }
        // Trim trailing whitespace per line
        let trimmed = row_text.trim_end();
        text.push_str(trimmed);
        text.push('\n');
    }

    drop(term);

    // Trim trailing empty lines
    let text = text.trim_end_matches('\n').to_string();

    if !text.is_empty() {
        cx.write_to_clipboard(ClipboardItem::new_string(text));
    }
}

/// Handle files dropped onto the terminal panel.
///
/// - 1 file → `vgrid peek "./filename" --headers`
/// - 2 files → `vgrid diff "./a" "./b" --key <guess>`
/// - 3+ files → insert paths one per line (no command)
///
/// Files are copied into the workspace root (collision-safe) so commands use
/// relative paths. A status message confirms what happened.
fn handle_terminal_drop(
    app: &mut Spreadsheet,
    paths: &[PathBuf],
    window: &mut Window,
    cx: &mut Context<Spreadsheet>,
) {
    if paths.is_empty() {
        return;
    }

    // Ensure terminal is alive
    if app.terminal.term.is_none() && !app.terminal.exited {
        app.spawn_terminal(window, cx);
    }
    if app.terminal.term.is_none() {
        return;
    }

    let workspace = app.terminal.workspace_root.clone();

    // Copy files into workspace and collect relative names
    let mut local_names: Vec<String> = Vec::new();
    let mut original_paths: Vec<PathBuf> = Vec::new();
    let mut renames: Vec<String> = Vec::new();
    let mut copy_count = 0u32;

    for path in paths {
        if let Some(ref ws) = workspace {
            if let Some(filename) = path.file_name() {
                let dest = collision_safe_dest(ws, filename);
                let dest_name = dest.file_name().unwrap().to_string_lossy().to_string();

                // Track renames
                let orig_name = filename.to_string_lossy().to_string();
                if dest_name != orig_name {
                    renames.push(format!("{} -> {}", orig_name, dest_name));
                }

                // Copy file if it's not already in the workspace
                if dest != *path {
                    if let Err(e) = std::fs::copy(path, &dest) {
                        eprintln!("Failed to copy {:?} to workspace: {}", path, e);
                        local_names.push(format!("\"{}\"", path.display()));
                        original_paths.push(path.clone());
                        continue;
                    }
                    copy_count += 1;
                }

                local_names.push(format!("\"./{}\"", dest_name));
                original_paths.push(path.clone());
                continue;
            }
        }
        // Fallback: use absolute path
        local_names.push(format!("\"{}\"", path.display()));
        original_paths.push(path.clone());
    }

    let command = match local_names.len() {
        1 => {
            format!("vgrid peek {} --headers", local_names[0])
        }
        2 => {
            let key_flag = guess_diff_key(&original_paths[0], &original_paths[1]);
            let left_stem = crate::diff_view::sanitize_stem(
                &original_paths[0].file_name().unwrap_or_default().to_string_lossy(),
            );
            let right_stem = crate::diff_view::sanitize_stem(
                &original_paths[1].file_name().unwrap_or_default().to_string_lossy(),
            );
            let output_file = format!("diff-{}_vs_{}.json", left_stem, right_stem);
            format!(
                "vgrid diff {} {} --key {} --output ./{} # After running: Command Palette → Open Diff Results",
                local_names[0], local_names[1], key_flag, output_file,
            )
        }
        _ => {
            local_names.join(" ")
        }
    };

    // Clear current line (Ctrl+U) before inserting, then write command (no trailing \n)
    app.terminal.write_to_pty(b"\x15");
    app.terminal.write_to_pty(command.as_bytes());

    // Status message
    let mut status = if copy_count > 0 {
        format!("Copied {} file{} to workspace", copy_count, if copy_count == 1 { "" } else { "s" })
    } else {
        format!("Dropped {} file{}", paths.len(), if paths.len() == 1 { "" } else { "s" })
    };
    if !renames.is_empty() {
        status.push_str(&format!(" (renamed: {})", renames.join(", ")));
    }
    if local_names.len() == 2 {
        status.push_str(". Run the command, then use Command Palette → Open Diff Results");
    }
    app.status_message = Some(status);

    cx.notify();
}

/// Find a collision-safe destination path in the workspace directory.
///
/// If `workspace/filename.csv` already exists, tries `filename (2).csv`,
/// `filename (3).csv`, etc., up to 99 attempts.
fn collision_safe_dest(workspace: &Path, filename: &std::ffi::OsStr) -> PathBuf {
    let dest = workspace.join(filename);
    if !dest.exists() {
        return dest;
    }

    let name = filename.to_string_lossy();
    let (stem, ext) = match name.rfind('.') {
        Some(i) => (&name[..i], Some(&name[i..])), // includes the dot
        None => (name.as_ref(), None),
    };

    for n in 2..=99 {
        let new_name = match ext {
            Some(ext) => format!("{} ({}){}", stem, n, ext),
            None => format!("{} ({})", stem, n),
        };
        let candidate = workspace.join(&new_name);
        if !candidate.exists() {
            return candidate;
        }
    }

    // Extremely unlikely: fall back to original (will overwrite)
    dest
}

/// Guess a `--key` column for `vgrid diff` from two CSV files.
///
/// Reads headers from both files, normalizes them (trim, unquote, case-fold,
/// spaces→underscores), finds the intersection, and checks for common
/// identifier column names. Returns the original header name (quoted if it
/// contains spaces) or a `<KEY>` placeholder.
fn guess_diff_key(left: &Path, right: &Path) -> String {
    let is_csv = |p: &Path| {
        p.extension()
            .and_then(|e| e.to_str())
            .map(|e| matches!(e.to_lowercase().as_str(), "csv" | "tsv" | "txt"))
            .unwrap_or(false)
    };

    if !is_csv(left) || !is_csv(right) {
        return "<KEY>".to_string();
    }

    /// A parsed header: the original text and a normalized form for matching.
    struct Header {
        original: String,
        normalized: String,
    }

    let read_headers = |p: &Path| -> Option<Vec<Header>> {
        let mut rdr = std::io::BufReader::new(std::fs::File::open(p).ok()?);
        let mut first_line = String::new();
        std::io::BufRead::read_line(&mut rdr, &mut first_line).ok()?;
        let trimmed = first_line.trim();
        if trimmed.is_empty() {
            return None;
        }
        let delim = if p.extension().and_then(|e| e.to_str()) == Some("tsv") {
            '\t'
        } else {
            ','
        };
        Some(
            trimmed
                .split(delim)
                .map(|h| {
                    let original = h.trim().trim_matches('"').to_string();
                    let normalized = original
                        .to_lowercase()
                        .replace([' ', '-'], "_");
                    Header { original, normalized }
                })
                .collect(),
        )
    };

    let left_headers = match read_headers(left) {
        Some(h) => h,
        None => return "<KEY>".to_string(),
    };
    let right_headers = match read_headers(right) {
        Some(h) => h,
        None => return "<KEY>".to_string(),
    };

    const KEY_CANDIDATES: &[&str] = &[
        "id", "uuid", "external_id", "transaction_id", "txn_id",
        "payout_id", "invoice_id", "order_id", "sku",
        "email", "name",
    ];

    // Build normalized→original maps, then find intersection by normalized key
    let left_map: std::collections::HashMap<&str, &str> =
        left_headers.iter().map(|h| (h.normalized.as_str(), h.original.as_str())).collect();
    let right_set: std::collections::HashSet<&str> =
        right_headers.iter().map(|h| h.normalized.as_str()).collect();

    for candidate in KEY_CANDIDATES {
        if left_map.contains_key(candidate) && right_set.contains(candidate) {
            let original = left_map[candidate];
            // Quote if the original name contains spaces
            if original.contains(' ') {
                return format!("\"{}\"", original);
            }
            return original.to_string();
        }
    }

    "<KEY>".to_string()
}

/// Map an ANSI color to GPUI Hsla.
fn ansi_to_hsla(color: AnsiColor, default_fg: Hsla, default_bg: Hsla) -> Hsla {
    match color {
        AnsiColor::Named(named) => named_color_to_hsla(named, default_fg, default_bg),
        AnsiColor::Spec(rgb) => {
            let r = rgb.r as f32 / 255.0;
            let g = rgb.g as f32 / 255.0;
            let b = rgb.b as f32 / 255.0;
            rgb_to_hsla(r, g, b)
        }
        AnsiColor::Indexed(idx) => indexed_color_to_hsla(idx),
    }
}

/// Resolve a cell's display character, handling spacers and hidden cells.
fn cell_char(c: char, flags: CellFlags) -> char {
    if flags.contains(CellFlags::WIDE_CHAR_SPACER)
        || flags.contains(CellFlags::LEADING_WIDE_CHAR_SPACER)
        || flags.contains(CellFlags::HIDDEN)
    {
        ' '
    } else {
        c
    }
}

/// Build a styled text span for a terminal run.
fn make_span(
    text: String,
    fg: AnsiColor,
    bg: AnsiColor,
    flags: CellFlags,
    text_primary: Hsla,
    editor_bg: Hsla,
) -> gpui::Div {
    let (effective_fg, effective_bg) = if flags.contains(CellFlags::INVERSE) {
        (bg, fg)
    } else {
        (fg, bg)
    };

    let fg_hsla = ansi_to_hsla(effective_fg, text_primary, editor_bg);
    let bg_hsla = ansi_to_hsla(effective_bg, text_primary, editor_bg);

    let mut span = div()
        .font_family(TERM_FONT_FAMILY)
        .text_color(fg_hsla)
        .text_size(px(TERM_FONT_SIZE));

    if !is_default_bg(effective_bg) {
        span = span.bg(bg_hsla);
    }
    if flags.contains(CellFlags::BOLD) {
        span = span.font_weight(FontWeight::BOLD);
    }
    if flags.contains(CellFlags::ITALIC) {
        span = span.italic();
    }
    if flags.contains(CellFlags::UNDERLINE) || flags.contains(CellFlags::DOUBLE_UNDERLINE) {
        span = span.underline();
    }
    if flags.contains(CellFlags::STRIKEOUT) {
        span = span.line_through();
    }

    span.child(text)
}

/// Check if a color is the default background.
fn is_default_bg(color: AnsiColor) -> bool {
    matches!(color, AnsiColor::Named(NamedColor::Background))
}

/// Map a named ANSI color to Hsla.
fn named_color_to_hsla(color: NamedColor, default_fg: Hsla, default_bg: Hsla) -> Hsla {
    match color {
        // Standard colors
        NamedColor::Black => rgb_to_hsla(0.0, 0.0, 0.0),
        NamedColor::Red => rgb_to_hsla(0.8, 0.0, 0.0),
        NamedColor::Green => rgb_to_hsla(0.18, 0.65, 0.18),
        NamedColor::Yellow => rgb_to_hsla(0.75, 0.65, 0.0),
        NamedColor::Blue => rgb_to_hsla(0.2, 0.4, 0.9),
        NamedColor::Magenta => rgb_to_hsla(0.7, 0.2, 0.7),
        NamedColor::Cyan => rgb_to_hsla(0.0, 0.65, 0.65),
        NamedColor::White => rgb_to_hsla(0.75, 0.75, 0.75),
        // Bright colors
        NamedColor::BrightBlack => rgb_to_hsla(0.5, 0.5, 0.5),
        NamedColor::BrightRed => rgb_to_hsla(1.0, 0.33, 0.33),
        NamedColor::BrightGreen => rgb_to_hsla(0.3, 0.8, 0.3),
        NamedColor::BrightYellow => rgb_to_hsla(0.85, 0.75, 0.1),
        NamedColor::BrightBlue => rgb_to_hsla(0.4, 0.55, 1.0),
        NamedColor::BrightMagenta => rgb_to_hsla(0.85, 0.4, 0.85),
        NamedColor::BrightCyan => rgb_to_hsla(0.2, 0.8, 0.8),
        NamedColor::BrightWhite => rgb_to_hsla(1.0, 1.0, 1.0),
        // Dim colors
        NamedColor::DimBlack => rgb_to_hsla(0.0, 0.0, 0.0),
        NamedColor::DimRed => rgb_to_hsla(0.55, 0.0, 0.0),
        NamedColor::DimGreen => rgb_to_hsla(0.0, 0.55, 0.0),
        NamedColor::DimYellow => rgb_to_hsla(0.55, 0.55, 0.0),
        NamedColor::DimBlue => rgb_to_hsla(0.0, 0.0, 0.55),
        NamedColor::DimMagenta => rgb_to_hsla(0.55, 0.0, 0.55),
        NamedColor::DimCyan => rgb_to_hsla(0.0, 0.55, 0.55),
        NamedColor::DimWhite => rgb_to_hsla(0.5, 0.5, 0.5),
        // Special
        NamedColor::Foreground | NamedColor::BrightForeground | NamedColor::DimForeground => default_fg,
        NamedColor::Background => default_bg,
        NamedColor::Cursor => default_fg,
    }
}

/// Map a 256-color index to Hsla.
fn indexed_color_to_hsla(idx: u8) -> Hsla {
    match idx {
        // Standard colors (0-7) — match named_color_to_hsla
        0 => rgb_to_hsla(0.0, 0.0, 0.0),
        1 => rgb_to_hsla(0.8, 0.0, 0.0),
        2 => rgb_to_hsla(0.18, 0.65, 0.18),
        3 => rgb_to_hsla(0.75, 0.65, 0.0),
        4 => rgb_to_hsla(0.2, 0.4, 0.9),
        5 => rgb_to_hsla(0.7, 0.2, 0.7),
        6 => rgb_to_hsla(0.0, 0.65, 0.65),
        7 => rgb_to_hsla(0.75, 0.75, 0.75),
        // Bright colors (8-15)
        8 => rgb_to_hsla(0.5, 0.5, 0.5),
        9 => rgb_to_hsla(1.0, 0.33, 0.33),
        10 => rgb_to_hsla(0.3, 0.8, 0.3),
        11 => rgb_to_hsla(0.85, 0.75, 0.1),
        12 => rgb_to_hsla(0.4, 0.55, 1.0),
        13 => rgb_to_hsla(0.85, 0.4, 0.85),
        14 => rgb_to_hsla(0.2, 0.8, 0.8),
        15 => rgb_to_hsla(1.0, 1.0, 1.0),
        // 6x6x6 color cube (16-231)
        16..=231 => {
            let idx = idx - 16;
            let b = (idx % 6) as f32 / 5.0;
            let g = ((idx / 6) % 6) as f32 / 5.0;
            let r = (idx / 36) as f32 / 5.0;
            rgb_to_hsla(r, g, b)
        }
        // Grayscale ramp (232-255)
        232..=255 => {
            let v = (idx - 232) as f32 / 23.0;
            rgb_to_hsla(v, v, v)
        }
    }
}

/// Convert linear RGB (0.0–1.0) to GPUI Hsla.
fn rgb_to_hsla(r: f32, g: f32, b: f32) -> Hsla {
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;

    if (max - min).abs() < f32::EPSILON {
        return hsla(0.0, 0.0, l, 1.0);
    }

    let d = max - min;
    let s = if l > 0.5 {
        d / (2.0 - max - min)
    } else {
        d / (max + min)
    };

    let h = if (max - r).abs() < f32::EPSILON {
        let mut h = (g - b) / d;
        if g < b {
            h += 6.0;
        }
        h / 6.0
    } else if (max - g).abs() < f32::EPSILON {
        ((b - r) / d + 2.0) / 6.0
    } else {
        ((r - g) / d + 4.0) / 6.0
    };

    hsla(h, s, l, 1.0)
}
