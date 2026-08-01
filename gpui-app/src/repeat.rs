//! Repeat last action — Excel's *other* F4.
//!
//! F4 in Excel means two different things depending on mode: cycle a
//! reference's absoluteness while editing a formula (which VisiGrid has
//! always had), and repeat the last command in normal mode (which it has
//! not). This is the second one.
//!
//! ## Where the slot is written, and why
//!
//! `set_repeat` is called inside the `Spreadsheet` mutation methods, not at
//! a dispatch layer. Every user path — keybinding, command palette, menu,
//! format bar, context menu — converges on those methods, so one line each
//! covers all of them; routing every `on_action` through `dispatch_command`
//! instead would mean rewriting ~50 handlers, several of which have no
//! `CommandId` at all.
//!
//! It also captures INTENT rather than effect. Deriving the slot from undo
//! history looks tempting — every mutation records one — but history stores
//! per-cell before/after patches, and toggles are ambiguous in that form:
//! Ctrl+B over a mixed-bold selection records some cells going true and
//! others staying, while the user's intent ("make them all bold") is a
//! single value that only the calling method knows.
//!
//! A useful side effect of this seam: agents cannot hijack what F4 repeats.
//! Session clients apply cell and format ops through the ENGINE's `Sheet`
//! (see `visigrid-session-host::handlers`), never through these methods, so
//! the slot always holds something the human did. The one exception is
//! structural ops, which the session adapter does route through the GUI's
//! `insert_rows`/`delete_rows`; that path suppresses the slot explicitly.

use gpui::*;
use visigrid_engine::cell::{Alignment, CellStyle, NumberFormat, VerticalAlignment};

use crate::app::Spreadsheet;
use crate::formatting::BorderApplyMode;

/// A command that F4 can re-apply to the current selection.
///
/// Deliberately excluded: text entry (Excel does not retype a value on F4),
/// navigation, selection, undo/redo itself, and file operations. Paste is
/// excluded for now — Excel does repeat it, but it interacts with clipboard
/// state in ways worth designing separately.
#[derive(Debug, Clone, PartialEq)]
pub enum RepeatAction {
    Bold(bool),
    Italic(bool),
    Underline(bool),
    Strikethrough(bool),
    Alignment(Alignment),
    VerticalAlignment(VerticalAlignment),
    NumberFormat(NumberFormat),
    CellStyle(CellStyle),
    BackgroundColor(Option<[u8; 4]>),
    FontColor(Option<[u8; 4]>),
    FontFamily(Option<String>),
    FontSize(Option<f32>),
    Borders(BorderApplyMode),
    ClearFormatting,
    MergeCells,
    UnmergeCells,
    InsertRows(usize),
    DeleteRows(usize),
    InsertCols(usize),
    DeleteCols(usize),
    FitColumnWidth,
}

impl RepeatAction {
    /// Menu-style label, shown in the status message when F4 fires.
    pub fn label(&self) -> String {
        match self {
            RepeatAction::Bold(v) => format!("{} bold", if *v { "Apply" } else { "Remove" }),
            RepeatAction::Italic(v) => format!("{} italic", if *v { "Apply" } else { "Remove" }),
            RepeatAction::Underline(v) => format!("{} underline", if *v { "Apply" } else { "Remove" }),
            RepeatAction::Strikethrough(v) => {
                format!("{} strikethrough", if *v { "Apply" } else { "Remove" })
            }
            RepeatAction::Alignment(_) => "Align".to_string(),
            RepeatAction::VerticalAlignment(_) => "Vertical align".to_string(),
            RepeatAction::NumberFormat(_) => "Number format".to_string(),
            RepeatAction::CellStyle(_) => "Cell style".to_string(),
            RepeatAction::BackgroundColor(_) => "Fill color".to_string(),
            RepeatAction::FontColor(_) => "Font color".to_string(),
            RepeatAction::FontFamily(_) => "Font".to_string(),
            RepeatAction::FontSize(_) => "Font size".to_string(),
            RepeatAction::Borders(_) => "Borders".to_string(),
            RepeatAction::ClearFormatting => "Clear formatting".to_string(),
            RepeatAction::MergeCells => "Merge cells".to_string(),
            RepeatAction::UnmergeCells => "Unmerge cells".to_string(),
            RepeatAction::InsertRows(n) => format!("Insert {} row(s)", n),
            RepeatAction::DeleteRows(n) => format!("Delete {} row(s)", n),
            RepeatAction::InsertCols(n) => format!("Insert {} column(s)", n),
            RepeatAction::DeleteCols(n) => format!("Delete {} column(s)", n),
            RepeatAction::FitColumnWidth => "Fit column to width".to_string(),
        }
    }
}

impl Spreadsheet {
    /// Record the last repeatable command. Called from the mutation methods
    /// themselves — see the module docs for why that is the right seam.
    pub(crate) fn set_repeat(&mut self, action: RepeatAction) {
        if self.suppress_repeat_capture {
            return;
        }
        self.repeat_action = Some(action);
    }

    /// What F4 would do right now, if anything.
    pub fn repeat_label(&self) -> Option<String> {
        self.repeat_action.as_ref().map(|a| a.label())
    }

    /// Re-apply the last repeatable command to the current selection.
    ///
    /// Typing a value neither repeats nor clears the slot: bold a cell, type
    /// in three others, press F4, and the bold still applies — matching
    /// Excel, where the slot holds the last *command*, not the last edit.
    pub fn repeat_last_action(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let Some(action) = self.repeat_action.clone() else {
            self.status_message = Some("Nothing to repeat".to_string());
            cx.notify();
            return;
        };
        if self.block_if_previewing(cx) {
            return;
        }

        // Re-applying must not overwrite the slot with itself — otherwise a
        // command whose parameters depend on the selection (structure ops)
        // would drift on each press.
        self.suppress_repeat_capture = true;

        let (row, col) = self.view_state.active_cell();
        match action.clone() {
            RepeatAction::Bold(v) => self.set_bold(v, cx),
            RepeatAction::Italic(v) => self.set_italic(v, cx),
            RepeatAction::Underline(v) => self.set_underline(v, cx),
            RepeatAction::Strikethrough(v) => self.set_strikethrough(v, cx),
            RepeatAction::Alignment(a) => self.set_alignment_selection(a, cx),
            RepeatAction::VerticalAlignment(a) => self.set_vertical_alignment_selection(a, cx),
            RepeatAction::NumberFormat(f) => self.set_number_format_selection(f, cx),
            RepeatAction::CellStyle(s) => self.set_cell_style_selection(s, cx),
            RepeatAction::BackgroundColor(c) => self.set_background_color(c, cx),
            RepeatAction::FontColor(c) => self.set_font_color_selection(c, cx),
            RepeatAction::FontFamily(f) => self.set_font_family_selection(f, cx),
            RepeatAction::FontSize(s) => self.set_font_size_selection(s, cx),
            RepeatAction::Borders(m) => self.apply_borders(m, cx),
            RepeatAction::ClearFormatting => self.clear_formatting_selection(cx),
            RepeatAction::MergeCells => self.merge_cells(cx),
            RepeatAction::UnmergeCells => self.unmerge_cells(cx),
            // Structure ops repeat at the CURRENT selection, not the original
            // one — "insert a row here too" is the whole point.
            RepeatAction::InsertRows(n) => self.insert_rows(row, n, cx),
            RepeatAction::DeleteRows(n) => self.delete_rows(row, n, cx),
            RepeatAction::InsertCols(n) => self.insert_cols(col, n, cx),
            RepeatAction::DeleteCols(n) => self.delete_cols(col, n, cx),
            RepeatAction::FitColumnWidth => self.fit_selection_columns(window, cx),
        }

        self.suppress_repeat_capture = false;
        self.status_message = Some(format!("Repeated: {}", action.label()));
        cx.notify();
    }
}

#[cfg(test)]
mod tests {
    // Explicit import, not `use super::*`: this module glob-imports gpui,
    // whose `test` attribute macro would shadow the built-in `#[test]`.
    use super::RepeatAction;

    #[test]
    fn labels_describe_the_command() {
        assert_eq!(RepeatAction::Bold(true).label(), "Apply bold");
        assert_eq!(RepeatAction::Bold(false).label(), "Remove bold");
        assert_eq!(RepeatAction::InsertRows(3).label(), "Insert 3 row(s)");
        assert_eq!(RepeatAction::ClearFormatting.label(), "Clear formatting");
    }

    #[test]
    fn toggles_carry_the_resolved_value_not_the_gesture() {
        // The point of capturing intent rather than effect: F4 after
        // Ctrl+B on a mixed selection must apply bold everywhere, so the
        // slot holds a concrete value, never "toggle".
        assert_ne!(RepeatAction::Bold(true), RepeatAction::Bold(false));
    }
}
