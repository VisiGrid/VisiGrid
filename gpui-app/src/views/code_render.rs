//! Shared syntax-highlighted code rendering helpers.
//!
//! Used by both `lua_console.rs` (REPL input area) and `script_view.rs` (script editor).

use std::ops::Range;

use gpui::*;
use gpui::prelude::FluentBuilder;

use crate::app::Spreadsheet;
use crate::scripting::LuaTokenType;
use crate::theme::TokenKey;

/// Return the color for a Lua token type, using the app's theme.
pub fn lua_token_color(app: &Spreadsheet, tt: LuaTokenType) -> Hsla {
    match tt {
        LuaTokenType::Keyword => app.token(TokenKey::FormulaFunction),
        LuaTokenType::Boolean => app.token(TokenKey::FormulaBoolean),
        LuaTokenType::String => app.token(TokenKey::FormulaString),
        LuaTokenType::Number => app.token(TokenKey::FormulaNumber),
        LuaTokenType::Comment => app.token(TokenKey::TextMuted),
        LuaTokenType::Operator | LuaTokenType::Punctuation => app.token(TokenKey::FormulaOperator),
        LuaTokenType::Identifier => app.token(TokenKey::TextPrimary),
    }
}

/// Render a single line with syntax-highlighted spans and optional cursor.
///
/// Uses binary search to skip tokens before this line, and processes tokens
/// inline without intermediate Vec allocations. For a file with N total tokens
/// this is O(log N + tokens_in_line) per line instead of O(N).
pub fn render_highlighted_line(
    line_text: &str,
    line_start: usize,
    line_end: usize,
    tokens: &[(Range<usize>, LuaTokenType)],
    cursor_in_line: Option<usize>,
    token_color: &dyn Fn(LuaTokenType) -> Hsla,
    text_primary: Hsla,
    accent: Hsla,
) -> Div {
    let mut row = div()
        .flex_1()
        .pl(px(6.0))
        .flex()
        .items_center()
        .overflow_hidden();

    let line_len = line_text.len();

    // Empty line: just show cursor if present
    if line_len == 0 {
        if cursor_in_line.is_some() {
            row = row.child(
                div().w(px(1.5)).h(px(12.0)).bg(accent).flex_shrink_0()
            );
        }
        return row;
    }

    // Binary search: find first token whose end > line_start (could overlap this line).
    let token_start = tokens.partition_point(|(range, _)| range.end <= line_start);

    // Walk tokens from token_start, emitting gap + token spans inline (no Vec needed).
    let mut pos: usize = 0; // position relative to line_start

    for (range, tt) in &tokens[token_start..] {
        if range.start >= line_end {
            break; // past this line, done
        }

        // Clip token to line bounds (relative to line_start)
        let clipped_start = range.start.max(line_start) - line_start;
        let clipped_end = range.end.min(line_end) - line_start;
        if clipped_start >= clipped_end {
            continue;
        }

        // Gap before this token
        if clipped_start > pos {
            row = emit_span(row, line_text, pos, clipped_start, text_primary, cursor_in_line, accent);
        }

        // The token itself
        row = emit_span(row, line_text, clipped_start, clipped_end, token_color(*tt), cursor_in_line, accent);
        pos = clipped_end;
    }

    // Trailing gap after last token
    if pos < line_len {
        row = emit_span(row, line_text, pos, line_len, text_primary, cursor_in_line, accent);
    }

    // Cursor at end of line (past all text)
    if let Some(cur) = cursor_in_line {
        if cur >= line_len {
            row = row.child(
                div().w(px(1.5)).h(px(12.0)).bg(accent).flex_shrink_0()
            );
        }
    }

    row
}

/// Emit a text span into the row, splitting at cursor position if needed.
#[inline]
fn emit_span(
    mut row: Div,
    line_text: &str,
    start: usize,
    end: usize,
    color: Hsla,
    cursor_in_line: Option<usize>,
    accent: Hsla,
) -> Div {
    if let Some(cur) = cursor_in_line {
        if cur >= start && cur < end {
            // Cursor inside this span: split into before | cursor | after
            if cur > start {
                row = row.child(div().text_color(color).child(line_text[start..cur].to_string()));
            }
            row = row.child(div().w(px(1.5)).h(px(12.0)).bg(accent).flex_shrink_0());
            if cur < end {
                row = row.child(div().text_color(color).child(line_text[cur..end].to_string()));
            }
            return row;
        }
    }
    row.child(div().text_color(color).child(line_text[start..end].to_string()))
}

/// Render a complete code line with gutter (line number) and highlighted code.
pub fn render_code_line(
    line_text: &str,
    line_start: usize,
    line_end: usize,
    tokens: &[(Range<usize>, LuaTokenType)],
    cursor_in_line: Option<usize>,
    line_num: usize,
    is_cursor_line: bool,
    gutter_width: f32,
    line_height: f32,
    text_primary: Hsla,
    text_muted: Hsla,
    accent: Hsla,
    token_color: &dyn Fn(LuaTokenType) -> Hsla,
) -> Div {
    div()
        .h(px(line_height))
        .flex()
        .items_center()
        .when(is_cursor_line, |d| d.bg(accent.opacity(0.10)))
        // Gutter: line number
        .child(
            div()
                .w(px(gutter_width))
                .flex_shrink_0()
                .flex()
                .justify_end()
                .pr(px(6.0))
                .text_size(px(9.0))
                .text_color(text_muted.opacity(0.55))
                .border_r_1()
                .border_color(text_muted.opacity(0.18))
                .h_full()
                .items_center()
                .child(format!("{}", line_num))
        )
        // Code area
        .child(
            render_highlighted_line(
                line_text, line_start, line_end,
                tokens, cursor_in_line,
                token_color, text_primary, accent,
            )
        )
}
