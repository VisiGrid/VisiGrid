//! Extract text from the terminal grid for structured result parsing.
//!
//! Extracts the last N lines from the terminal grid (visible viewport + scrollback).
//! Does not extract full scrollback — capped at `max_lines` (default 2000).

use std::sync::Arc;

use alacritty_terminal::grid::Dimensions;
use alacritty_terminal::index::{Column, Line};
use alacritty_terminal::sync::FairMutex;
use alacritty_terminal::term::Term;
use alacritty_terminal::term::cell::Flags as CellFlags;

use super::pty::TerminalEventProxy;

/// Extract the last `max_lines` lines of text from the terminal grid.
///
/// Walks backwards from the current bottom of the visible viewport,
/// including scrollback up to `max_lines`. Trims trailing whitespace per line.
/// Skips wide-char spacers (same pattern as `copy_viewport_text`).
pub fn extract_recent_text(
    term: &Arc<FairMutex<Term<TerminalEventProxy>>>,
    max_lines: usize,
) -> String {
    let term = term.lock();
    let cols = term.columns();
    let screen_lines = term.screen_lines();
    let total_lines = term.total_lines();

    // We want the last `max_lines` lines of total content.
    // total_lines = scrollback_lines + screen_lines.
    // In the alacritty grid, Line(0) is the top of the visible screen,
    // and negative line indices go into scrollback.
    // The bottommost visible line is Line(screen_lines - 1).
    // Scrollback lines are Line(-1), Line(-2), etc.

    let lines_to_extract = max_lines.min(total_lines);
    if lines_to_extract == 0 {
        return String::new();
    }

    // Start from the bottommost visible line and work up
    // Line indices: screen bottom = screen_lines - 1, scrollback = -1, -2, ...
    // We want the last `lines_to_extract` lines.
    // The bottom of the grid is Line(screen_lines - 1).
    // Going `lines_to_extract` up from there:
    //   start_line = (screen_lines - 1) - (lines_to_extract - 1) = screen_lines - lines_to_extract
    // This can be negative if we're going into scrollback.

    let start_line = screen_lines as i32 - lines_to_extract as i32;

    let mut lines: Vec<String> = Vec::with_capacity(lines_to_extract);

    for i in 0..lines_to_extract {
        let line_idx = start_line + i as i32;
        let line = Line(line_idx);
        let row = &term.grid()[line];
        let mut row_text = String::with_capacity(cols);
        for col_idx in 0..cols {
            let cell = &row[Column(col_idx)];
            if cell.flags.contains(CellFlags::WIDE_CHAR_SPACER)
                || cell.flags.contains(CellFlags::LEADING_WIDE_CHAR_SPACER)
            {
                continue;
            }
            row_text.push(cell.c);
        }
        lines.push(row_text.trim_end().to_string());
    }

    drop(term);

    // Trim trailing empty lines
    while lines.last().map_or(false, |l| l.is_empty()) {
        lines.pop();
    }

    lines.join("\n")
}
