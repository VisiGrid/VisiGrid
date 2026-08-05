//! Cell formatting methods for Spreadsheet
//!
//! This module contains all format setters (bold, italic, alignment, etc.)
//! and merge/unmerge cell operations.

use gpui::*;
use visigrid_engine::cell::{max_border, Alignment, CellBorder, CellFormat, CellStyle, NumberFormat, TextOverflow, VerticalAlignment};
use visigrid_engine::sheet::MergedRegion;

use crate::app::{Spreadsheet, TriState, SelectionFormatState, NUM_ROWS, NUM_COLS};
use crate::history::{CellFormatPatch, FormatActionKind, UndoAction};
use crate::mode::Mode;
use crate::repeat::RepeatAction;

/// Format Painter state: captured format + locked flag.
#[derive(Debug, Clone)]
pub struct FormatPaintState {
    /// The captured cell format snapshot (immutable until next CopyFormat).
    pub snapshot: CellFormat,
    /// If true, painter stays active after each apply (double-click / Shift+click).
    pub locked: bool,
}

/// Border application mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderApplyMode {
    /// Apply thin black borders to all 4 edges of each selected cell
    All,
    /// Apply thin black borders only to the outer perimeter of the selection
    Outline,
    /// Apply thin black borders only to internal edges (not outer perimeter)
    Inside,
    /// Apply thin black border to the top edge of the selection
    Top,
    /// Apply thin black border to the bottom edge of the selection
    Bottom,
    /// Apply thin black border to the left edge of the selection
    Left,
    /// Apply thin black border to the right edge of the selection
    Right,
    /// Clear all borders from selected cells
    Clear,
}

/// Maximum number of cells to scan when resolving selection format state.
/// This runs on every render of the format bar, so an uncapped scan freezes
/// the UI for seconds on huge selections (Ctrl+A = 16.7M cells). Mirrors
/// MAX_STATS_CELLS in status_bar.rs. Beyond the cap, tri-state resolution is
/// based on the first N cells of the selection (which includes the viewport
/// for a fresh select-all).
const MAX_FORMAT_STATE_CELLS: usize = 10_000;

/// Above this many selected cells, a formatting command is clamped to the
/// cells that actually hold something. See `format_apply_ranges`.
///
/// Deliberately the same number as the read-side cap. A first draft used
/// 100_000, which looked generous until a test pointed out that one whole
/// column is 65_536 cells — so clicking a column header and pressing Ctrl+B,
/// an everyday gesture, still sat under the cap and materialised 65_536 cells
/// plus 65_536 undo patches.
const MAX_FORMAT_APPLY_CELLS: usize = MAX_FORMAT_STATE_CELLS;

impl Spreadsheet {
    /// The bounding box of every populated cell, or None on an empty sheet.
    ///
    /// Cheap regardless of grid size: the cell map is sparse, so this walks
    /// what exists, not the 16.7M coordinates that could exist.
    fn populated_bounds(&self, cx: &App) -> Option<(usize, usize, usize, usize)> {
        let mut bounds: Option<(usize, usize, usize, usize)> = None;
        for (&(row, col), _) in self.sheet(cx).cells_iter() {
            bounds = Some(match bounds {
                None => (row, col, row, col),
                Some((r0, c0, r1, c1)) => (r0.min(row), c0.min(col), r1.max(row), c1.max(col)),
            });
        }
        bounds
    }

    /// The ranges a formatting command should actually write to.
    ///
    /// Formatting lives on the cell, and the cell map is sparse — so writing
    /// a format to a coordinate MATERIALISES a cell there. Ctrl+A then Ctrl+B
    /// meant inserting 16.7M cells, cloning two `CellFormat`s per cell into an
    /// undo patch, and then saving all of it to disk. The app hung, and the
    /// file would have grown by orders of magnitude. This is the write-side
    /// counterpart to `MAX_FORMAT_STATE_CELLS`, which capped the read-side
    /// scan for the same reason (issue #5).
    ///
    /// Selections below the cap are applied exactly as given, so the ordinary
    /// case is untouched: format an empty block, type into it, and the text is
    /// still formatted. Past the cap — select-all, or a column header click —
    /// the clamped result is visually identical anyway, because a format on a
    /// cell with no content and no neighbours renders nothing.
    ///
    /// The limitation this accepts: after a clamped apply, typing into a cell
    /// that was outside the populated area does not inherit the format.
    /// Excel would inherit it, because Excel stores row/column/sheet-level
    /// defaults. Doing that properly needs a format hierarchy in the engine,
    /// not a bigger loop here.
    pub(crate) fn format_apply_ranges(&self, cx: &App) -> Vec<((usize, usize), (usize, usize))> {
        // The sparse scan is lazy: it only runs when the selection is over the
        // cap, so the ordinary path costs nothing.
        plan_format_ranges(self.all_selection_ranges(), || self.populated_bounds(cx))
    }
    /// Compute format state for the current selection (tri-state resolution)
    pub fn selection_format_state(&self, cx: &App) -> SelectionFormatState {
        let mut state = SelectionFormatState::default();
        let mut first = true;
        let mut last_display: Option<String> = None;
        let mut scanned = 0usize;

        let ranges = self.all_selection_ranges();
        // True cell count, independent of the scan cap below
        state.cell_count = ranges
            .iter()
            .map(|((r1, c1), (r2, c2))| (r2 - r1 + 1) * (c2 - c1 + 1))
            .sum();

        'scan: for ((min_row, min_col), (max_row, max_col)) in ranges {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    if scanned >= MAX_FORMAT_STATE_CELLS {
                        break 'scan;
                    }
                    scanned += 1;
                    let raw = self.sheet(cx).get_raw(row, col);
                    let display = self.sheet(cx).get_display(row, col);
                    let format = self.sheet(cx).get_format(row, col);

                    if first {
                        state.raw_value = TriState::Uniform(raw.clone());
                        state.bold = TriState::Uniform(format.bold);
                        state.italic = TriState::Uniform(format.italic);
                        state.underline = TriState::Uniform(format.underline);
                        state.strikethrough = TriState::Uniform(format.strikethrough);
                        state.font_family = TriState::Uniform(format.font_family.clone());
                        state.alignment = TriState::Uniform(format.alignment);
                        state.vertical_alignment = TriState::Uniform(format.vertical_alignment);
                        state.text_overflow = TriState::Uniform(format.text_overflow);
                        state.number_format = TriState::Uniform(format.number_format);
                        state.background_color = TriState::Uniform(format.background_color);
                        state.font_size = TriState::Uniform(format.font_size);
                        state.font_color = TriState::Uniform(format.font_color);
                        state.cell_style = TriState::Uniform(format.cell_style);
                        last_display = Some(display);
                        first = false;
                    } else {
                        state.raw_value = state.raw_value.combine(&raw);
                        state.bold = state.bold.combine(&format.bold);
                        state.italic = state.italic.combine(&format.italic);
                        state.underline = state.underline.combine(&format.underline);
                        state.strikethrough = state.strikethrough.combine(&format.strikethrough);
                        state.font_family = state.font_family.combine(&format.font_family);
                        state.alignment = state.alignment.combine(&format.alignment);
                        state.vertical_alignment = state.vertical_alignment.combine(&format.vertical_alignment);
                        state.text_overflow = state.text_overflow.combine(&format.text_overflow);
                        state.number_format = state.number_format.combine(&format.number_format);
                        state.background_color = state.background_color.combine(&format.background_color);
                        state.font_size = state.font_size.combine(&format.font_size);
                        state.font_color = state.font_color.combine(&format.font_color);
                        state.cell_style = state.cell_style.combine(&format.cell_style);
                        last_display = Some(display);

                        // Every property already resolved to Mixed — scanning
                        // further cells cannot change the outcome.
                        if state.raw_value.is_mixed()
                            && state.bold.is_mixed()
                            && state.italic.is_mixed()
                            && state.underline.is_mixed()
                            && state.strikethrough.is_mixed()
                            && state.font_family.is_mixed()
                            && state.alignment.is_mixed()
                            && state.vertical_alignment.is_mixed()
                            && state.text_overflow.is_mixed()
                            && state.number_format.is_mixed()
                            && state.background_color.is_mixed()
                            && state.font_size.is_mixed()
                            && state.font_color.is_mixed()
                            && state.cell_style.is_mixed()
                        {
                            break 'scan;
                        }
                    }
                }
            }
        }

        // For single cell, show display value and extract numeric preview
        if matches!(state.raw_value, TriState::Uniform(_)) {
            state.display_value = last_display;
        }

        // Extract active cell numeric value for format preview
        if state.cell_count == 1 {
            let (row, col) = self.view_state.active_cell();
            if let visigrid_engine::cell::CellValue::Number(n) = self.sheet(cx).get_cell(row, col).value {
                state.preview_value = Some(n);
            }
        }

        state
    }

    /// Set bold on all selected cells (explicit value, not toggle)
    pub fn set_bold(&mut self, value: bool, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::Bold(value));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_bold(row, col, value));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = format!("Bold {}", if value { "on" } else { "off" });
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Bold, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set italic on all selected cells (explicit value, not toggle)
    pub fn set_italic(&mut self, value: bool, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::Italic(value));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_italic(row, col, value));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = format!("Italic {}", if value { "on" } else { "off" });
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Italic, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set underline on all selected cells (explicit value, not toggle)
    pub fn set_underline(&mut self, value: bool, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::Underline(value));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_underline(row, col, value));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = format!("Underline {}", if value { "on" } else { "off" });
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Underline, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set strikethrough on all selected cells (explicit value, not toggle)
    pub fn set_strikethrough(&mut self, value: bool, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::Strikethrough(value));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_strikethrough(row, col, value));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = format!("Strikethrough {}", if value { "on" } else { "off" });
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Strikethrough, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set font family on all selected cells
    pub fn set_font_family_selection(&mut self, font: Option<String>, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::FontFamily(font.clone()));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_font_family(row, col, font.clone()));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let font_name = font.as_deref().unwrap_or("default");
            let desc = format!("Font '{}'", font_name);
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Font, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set horizontal alignment on all selected cells
    pub fn set_alignment_selection(&mut self, alignment: Alignment, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::Alignment(alignment));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_alignment(row, col, alignment));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let align_name = match alignment {
                Alignment::General => "General",
                Alignment::Left => "Left",
                Alignment::Center => "Center",
                Alignment::Right => "Right",
                Alignment::CenterAcrossSelection => "Center Across",
            };
            let desc = format!("Align {}", align_name);
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Alignment, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Center Across Selection, toggle-style: if every selected cell is
    /// already CenterAcrossSelection, revert to General; otherwise apply.
    /// The merge-free alternative to Merge & Center — sorting, filtering,
    /// and formulas keep working because no cells are actually merged.
    pub fn center_across_selection_toggle(&mut self, cx: &mut Context<Self>) {
        let mut all_cas = true;
        'outer: for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    if self.sheet(cx).get_format(row, col).alignment != Alignment::CenterAcrossSelection {
                        all_cas = false;
                        break 'outer;
                    }
                }
            }
        }
        let target = if all_cas { Alignment::General } else { Alignment::CenterAcrossSelection };
        self.set_alignment_selection(target, cx);
    }

    /// Set vertical alignment on all selected cells
    pub fn set_vertical_alignment_selection(&mut self, valign: VerticalAlignment, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::VerticalAlignment(valign));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_vertical_alignment(row, col, valign));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let valign_name = match valign {
                VerticalAlignment::Top => "Top",
                VerticalAlignment::Middle => "Middle",
                VerticalAlignment::Bottom => "Bottom",
            };
            let desc = format!("V-Align {}", valign_name);
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::VerticalAlignment, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set text overflow on all selected cells
    pub fn set_text_overflow_selection(&mut self, overflow: TextOverflow, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_text_overflow(row, col, overflow));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let overflow_name = match overflow {
                TextOverflow::Clip => "Clip",
                TextOverflow::Wrap => "Wrap",
                TextOverflow::Overflow => "Overflow",
            };
            let desc = overflow_name.to_string();
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::TextOverflow, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set number format on all selected cells
    pub fn set_number_format_selection(&mut self, format: NumberFormat, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::NumberFormat(format.clone()));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    // Safety net: convert text "X%" to number when applying Percent format
                    if matches!(format, NumberFormat::Percent { .. }) {
                        let raw = self.sheet(cx).get_raw(row, col);
                        if let Some(pct) = raw.strip_suffix('%') {
                            let clean: String = pct.chars()
                                .filter(|c| !c.is_whitespace() && *c != ',')
                                .collect();
                            if let Ok(n) = clean.parse::<f64>() {
                                self.set_cell_value(row, col, &(n / 100.0).to_string(), cx);
                            }
                        }
                    }
                    let before = self.sheet(cx).get_format(row, col);
                    let fmt = format.clone();
                    self.active_sheet_mut(cx, |s| s.set_number_format(row, col, fmt));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let format_name = match &format {
                NumberFormat::General => "General",
                NumberFormat::Number { .. } => "Number",
                NumberFormat::Currency { .. } => "Currency",
                NumberFormat::Percent { .. } => "Percent",
                NumberFormat::Date { .. } => "Date",
                NumberFormat::Time => "Time",
                NumberFormat::DateTime => "DateTime",
                NumberFormat::Custom(_) => "Custom",
            };
            let desc = format!("{} format", format_name);
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::NumberFormat, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Open the number format editor dialog, populated from the active cell
    pub fn open_number_format_editor(&mut self, cx: &mut Context<Self>) {
        use crate::app::NumberFormatEditorState;
        use visigrid_engine::cell::CellValue;

        let (row, col) = self.view_state.active_cell();
        let format = self.sheet(cx).get_format(row, col);
        let cell = self.sheet(cx).get_cell(row, col);
        let sample = match cell.value {
            CellValue::Number(n) => n.abs(),
            _ => 1234.5678,
        };
        self.number_format_editor = NumberFormatEditorState::from_number_format(&format.number_format, sample);
        self.mode = Mode::NumberFormatEditor;
        cx.notify();
    }

    /// Close the number format editor without applying
    pub fn close_number_format_editor(&mut self, cx: &mut Context<Self>) {
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Apply the number format editor settings and close
    pub fn apply_number_format_editor(&mut self, cx: &mut Context<Self>) {
        let fmt = self.number_format_editor.to_number_format();
        self.set_number_format_selection(fmt, cx);
        self.mode = Mode::Navigation;
        cx.notify();
    }

    /// Adjust decimal places on selected cells - uses DecimalPlaces kind for coalescing
    pub fn adjust_decimals_selection(&mut self, delta: i8, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    let new_format = match &before.number_format {
                        NumberFormat::Number { decimals, thousands, negative } => {
                            let new_dec = (*decimals as i8 + delta).clamp(0, 10) as u8;
                            Some(NumberFormat::Number { decimals: new_dec, thousands: *thousands, negative: *negative })
                        }
                        NumberFormat::Currency { decimals, thousands, negative, symbol } => {
                            let new_dec = (*decimals as i8 + delta).clamp(0, 10) as u8;
                            Some(NumberFormat::Currency { decimals: new_dec, thousands: *thousands, negative: *negative, symbol: symbol.clone() })
                        }
                        NumberFormat::Percent { decimals } => {
                            let new_dec = (*decimals as i8 + delta).clamp(0, 10) as u8;
                            Some(NumberFormat::Percent { decimals: new_dec })
                        }
                        _ => None,
                    };
                    if let Some(fmt) = new_format {
                        self.active_sheet_mut(cx, |s| s.set_number_format(row, col, fmt));
                        let after = self.sheet(cx).get_format(row, col);
                        if before != after {
                            patches.push(CellFormatPatch { row, col, before, after });
                        }
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = format!("Decimal {}", if delta > 0 { "+" } else { "-" });
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::DecimalPlaces, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set background color on all selected cells
    pub fn set_background_color(&mut self, color: Option<[u8; 4]>, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::BackgroundColor(color));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_background_color(row, col, color));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = if color.is_some() { "Background color" } else { "Clear background" };
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::BackgroundColor, desc.to_string());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set font size on all selected cells
    pub fn set_font_size_selection(&mut self, size: Option<f32>, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::FontSize(size));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_font_size(row, col, size));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = if let Some(s) = size {
                format!("Font size {}", s as u32)
            } else {
                "Clear font size".to_string()
            };
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::FontSize, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set font color on all selected cells
    pub fn set_font_color_selection(&mut self, color: Option<[u8; 4]>, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::FontColor(color));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_font_color(row, col, color));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = if color.is_some() { "Text color" } else { "Clear text color" };
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::FontColor, desc.to_string());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    pub fn set_cell_style_selection(&mut self, style: CellStyle, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::CellStyle(style));
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| {
                        s.set_cell_style(row, col, style);
                        // Clear manual background/font color so style colors aren't hidden
                        if !style.is_none() {
                            s.set_background_color(row, col, None);
                            s.set_font_color(row, col, None);
                        }
                    });
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let desc = if style.is_none() {
                "Clear cell style".to_string()
            } else {
                format!("Cell Style: {}", style.label())
            };
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::CellStyle, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Start Format Painter (single-shot): capture the active cell's format.
    pub fn start_format_painter(&mut self, cx: &mut Context<Self>) {
        self.start_format_painter_inner(false, cx);
    }

    /// Start Format Painter in locked mode: stays active until Esc.
    pub fn start_format_painter_locked(&mut self, cx: &mut Context<Self>) {
        self.start_format_painter_inner(true, cx);
    }

    fn start_format_painter_inner(&mut self, locked: bool, cx: &mut Context<Self>) {
        let (row, col) = self.view_state.selected;
        let snapshot = self.sheet(cx).get_format(row, col);
        self.format_painter = Some(FormatPaintState { snapshot, locked });
        self.mode = crate::mode::Mode::FormatPainter;
        if locked {
            self.status_message = Some("Format Painter: LOCKED (Esc to cancel)".to_string());
        } else {
            self.status_message = Some("Format Painter: ON (Esc to cancel)".to_string());
        }
        cx.notify();
    }

    /// Copy format from active cell without entering FormatPainter mode (Ctrl+Shift+C).
    pub fn copy_format(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.view_state.selected;
        let snapshot = self.sheet(cx).get_format(row, col);
        self.format_painter = Some(FormatPaintState { snapshot, locked: false });
        self.status_message = Some("Format copied \u{00b7} Ctrl+Shift+V to paste".to_string());
        cx.notify();
    }

    /// Paste previously copied format onto current selection (Ctrl+Shift+V).
    pub fn paste_format(&mut self, cx: &mut Context<Self>) {
        let snapshot = match &self.format_painter {
            Some(state) => state.snapshot.clone(),
            None => {
                self.status_message = Some("No format copied \u{00b7} Ctrl+Shift+C first".to_string());
                cx.notify();
                return;
            }
        };
        // Paste format does not enter/exit FormatPainter mode — it's a one-shot apply
        self.apply_format_to_selection(&snapshot, cx);
        // Clear the captured format after paste (single-shot behavior)
        self.format_painter = None;
        cx.notify();
    }

    /// Apply Format Painter: set captured format on current selection.
    pub fn apply_format_painter(&mut self, cx: &mut Context<Self>) {
        let (snapshot, locked) = match &self.format_painter {
            Some(state) => (state.snapshot.clone(), state.locked),
            None => return,
        };

        self.apply_format_to_selection(&snapshot, cx);

        if locked {
            // Stay in FormatPainter mode — don't clear state
            self.status_message = Some("Format Painter: LOCKED (Esc to cancel)".to_string());
        } else {
            // Single-shot: disarm
            self.format_painter = None;
            self.mode = crate::mode::Mode::Navigation;
        }
        cx.notify();
    }

    /// Shared helper: apply a format snapshot to all selected cells with undo.
    fn apply_format_to_selection(&mut self, format: &CellFormat, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    if before != *format {
                        self.active_sheet_mut(cx, |s| s.set_format(row, col, format.clone()));
                        let after = self.sheet(cx).get_format(row, col);
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::PasteFormats, "Format Painter".to_string());
            self.is_modified = true;
            self.status_message = Some(format!("Format Painter → {} cell{}", count, if count == 1 { "" } else { "s" }));
        } else {
            self.status_message = None;
        }
    }

    /// Cancel Format Painter mode.
    pub fn cancel_format_painter(&mut self, cx: &mut Context<Self>) {
        self.format_painter = None;
        self.mode = crate::mode::Mode::Navigation;
        self.status_message = None;
        cx.notify();
    }

    /// Clear all formatting on selected cells, resetting to CellFormat::default().
    /// Records a single undo step regardless of cell count.
    pub fn clear_formatting_selection(&mut self, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::ClearFormatting);
        let mut patches = Vec::new();
        let default = CellFormat::default();
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    if before != default {
                        self.active_sheet_mut(cx, |s| s.set_format(row, col, default.clone()));
                        let after = self.sheet(cx).get_format(row, col);
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            // Rescan border flag: clearing formats may have removed the only bordered cells
            self.active_sheet_mut(cx, |s| s.scan_border_flag());
            let desc = "Clear Formatting".to_string();
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::ClearFormatting, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Apply borders to all selected cells with canonicalization.
    ///
    /// Canonicalization: UI commands set BOTH sides of every shared edge they touch
    /// to prevent conflicting border states from normal use.
    pub fn apply_borders(&mut self, mode: BorderApplyMode, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::Borders(mode));
        // Use current_border_color if set, otherwise None (Automatic = theme default)
        let thin = CellBorder {
            style: visigrid_engine::cell::BorderStyle::Thin,
            color: self.current_border_color,
        };
        let none = CellBorder::default();
        let mut patches = Vec::new();

        // For each selection range, apply borders with proper canonicalization
        for ((min_row, min_col), (max_row, max_col)) in self.format_apply_ranges(cx) {
            match mode {
                BorderApplyMode::All => {
                    // Set all 4 edges on each cell to Thin
                    for row in min_row..=max_row {
                        for col in min_col..=max_col {
                            let before = self.sheet(cx).get_format(row, col);
                            self.active_sheet_mut(cx, |s| s.set_borders(row, col, thin, thin, thin, thin));
                            let after = self.sheet(cx).get_format(row, col);
                            if before != after {
                                patches.push(CellFormatPatch { row, col, before, after });
                            }
                        }
                    }
                    // Internal edges are already consistent since we set all 4 on every cell
                }
                BorderApplyMode::Outline => {
                    // Set only perimeter edges, leave interior unchanged
                    for row in min_row..=max_row {
                        for col in min_col..=max_col {
                            let before = self.sheet(cx).get_format(row, col);
                            let mut changed = false;

                            // Top edge: only if on top row of selection
                            if row == min_row {
                                self.active_sheet_mut(cx, |s| s.set_border_top(row, col, thin));
                                changed = true;
                            }
                            // Bottom edge: only if on bottom row of selection
                            if row == max_row {
                                self.active_sheet_mut(cx, |s| s.set_border_bottom(row, col, thin));
                                changed = true;
                            }
                            // Left edge: only if on left column of selection
                            if col == min_col {
                                self.active_sheet_mut(cx, |s| s.set_border_left(row, col, thin));
                                changed = true;
                            }
                            // Right edge: only if on right column of selection
                            if col == max_col {
                                self.active_sheet_mut(cx, |s| s.set_border_right(row, col, thin));
                                changed = true;
                            }

                            if changed {
                                let after = self.sheet(cx).get_format(row, col);
                                if before != after {
                                    patches.push(CellFormatPatch { row, col, before, after });
                                }
                            }
                        }
                    }
                }
                BorderApplyMode::Inside => {
                    // Internal edges only: vertical internals as right edges,
                    // horizontal internals as bottom edges (precedence-aligned).
                    for row in min_row..=max_row {
                        for col in min_col..=max_col {
                            let is_internal_h = row < max_row;
                            let is_internal_v = col < max_col;
                            if !is_internal_h && !is_internal_v {
                                continue;
                            }
                            let before = self.sheet(cx).get_format(row, col);
                            if is_internal_h {
                                self.active_sheet_mut(cx, |s| s.set_border_bottom(row, col, thin));
                            }
                            if is_internal_v {
                                self.active_sheet_mut(cx, |s| s.set_border_right(row, col, thin));
                            }
                            let after = self.sheet(cx).get_format(row, col);
                            if before != after {
                                patches.push(CellFormatPatch { row, col, before, after });
                            }
                        }
                    }
                }
                BorderApplyMode::Top => {
                    // Top edge of selection: set top border on cells in min_row
                    for col in min_col..=max_col {
                        let before = self.sheet(cx).get_format(min_row, col);
                        self.active_sheet_mut(cx, |s| s.set_border_top(min_row, col, thin));
                        let after = self.sheet(cx).get_format(min_row, col);
                        if before != after {
                            patches.push(CellFormatPatch { row: min_row, col, before, after });
                        }
                    }
                }
                BorderApplyMode::Bottom => {
                    // Bottom edge of selection: set bottom border on cells in max_row
                    for col in min_col..=max_col {
                        let before = self.sheet(cx).get_format(max_row, col);
                        self.active_sheet_mut(cx, |s| s.set_border_bottom(max_row, col, thin));
                        let after = self.sheet(cx).get_format(max_row, col);
                        if before != after {
                            patches.push(CellFormatPatch { row: max_row, col, before, after });
                        }
                    }
                }
                BorderApplyMode::Left => {
                    // Left edge of selection: set left border on cells in min_col
                    for row in min_row..=max_row {
                        let before = self.sheet(cx).get_format(row, min_col);
                        self.active_sheet_mut(cx, |s| s.set_border_left(row, min_col, thin));
                        let after = self.sheet(cx).get_format(row, min_col);
                        if before != after {
                            patches.push(CellFormatPatch { row, col: min_col, before, after });
                        }
                    }
                }
                BorderApplyMode::Right => {
                    // Right edge of selection: set right border on cells in max_col
                    for row in min_row..=max_row {
                        let before = self.sheet(cx).get_format(row, max_col);
                        self.active_sheet_mut(cx, |s| s.set_border_right(row, max_col, thin));
                        let after = self.sheet(cx).get_format(row, max_col);
                        if before != after {
                            patches.push(CellFormatPatch { row, col: max_col, before, after });
                        }
                    }
                }
                BorderApplyMode::Clear => {
                    // Clear all 4 edges on each cell
                    for row in min_row..=max_row {
                        for col in min_col..=max_col {
                            let before = self.sheet(cx).get_format(row, col);
                            self.active_sheet_mut(cx, |s| s.set_borders(row, col, none, none, none, none));
                            let after = self.sheet(cx).get_format(row, col);
                            if before != after {
                                patches.push(CellFormatPatch { row, col, before, after });
                            }
                        }
                    }

                    // Also clear adjacent cells' inward-facing edges (canonicalization)
                    // Clear top edge of cells above the selection
                    if min_row > 0 {
                        for col in min_col..=max_col {
                            let adj_row = min_row - 1;
                            let before = self.sheet(cx).get_format(adj_row, col);
                            self.active_sheet_mut(cx, |s| s.set_border_bottom(adj_row, col, none));
                            let after = self.sheet(cx).get_format(adj_row, col);
                            if before != after {
                                patches.push(CellFormatPatch { row: adj_row, col, before, after });
                            }
                        }
                    }
                    // Clear bottom edge of cells below the selection
                    if max_row + 1 < self.sheet(cx).rows {
                        for col in min_col..=max_col {
                            let adj_row = max_row + 1;
                            let before = self.sheet(cx).get_format(adj_row, col);
                            self.active_sheet_mut(cx, |s| s.set_border_top(adj_row, col, none));
                            let after = self.sheet(cx).get_format(adj_row, col);
                            if before != after {
                                patches.push(CellFormatPatch { row: adj_row, col, before, after });
                            }
                        }
                    }
                    // Clear right edge of cells to the left of the selection
                    if min_col > 0 {
                        for row in min_row..=max_row {
                            let adj_col = min_col - 1;
                            let before = self.sheet(cx).get_format(row, adj_col);
                            self.active_sheet_mut(cx, |s| s.set_border_right(row, adj_col, none));
                            let after = self.sheet(cx).get_format(row, adj_col);
                            if before != after {
                                patches.push(CellFormatPatch { row, col: adj_col, before, after });
                            }
                        }
                    }
                    // Clear left edge of cells to the right of the selection
                    if max_col + 1 < self.sheet(cx).cols {
                        for row in min_row..=max_row {
                            let adj_col = max_col + 1;
                            let before = self.sheet(cx).get_format(row, adj_col);
                            self.active_sheet_mut(cx, |s| s.set_border_left(row, adj_col, none));
                            let after = self.sheet(cx).get_format(row, adj_col);
                            if before != after {
                                patches.push(CellFormatPatch { row, col: adj_col, before, after });
                            }
                        }
                    }
                }
            }
        }

        // Rescan border flag after clearing: may have removed the only bordered cells
        if matches!(mode, BorderApplyMode::Clear) && !patches.is_empty() {
            self.active_sheet_mut(cx, |s| s.scan_border_flag());
        }

        let count = patches.len();
        if count > 0 {
            let desc = match mode {
                BorderApplyMode::All => "All borders",
                BorderApplyMode::Outline => "Outline",
                BorderApplyMode::Inside => "Inside borders",
                BorderApplyMode::Top => "Top border",
                BorderApplyMode::Bottom => "Bottom border",
                BorderApplyMode::Left => "Left border",
                BorderApplyMode::Right => "Right border",
                BorderApplyMode::Clear => "Clear borders",
            };
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Border, desc.to_string());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    // ── Merge / Unmerge ──────────────────────────────────────────────

    /// Merge selected cells into one. Shows data-loss dialog if non-origin cells have data.
    pub fn merge_cells(&mut self, cx: &mut Context<Self>) {
        // Guard: multi-selection not supported
        if !self.view_state.additional_selections.is_empty() {
            self.status_message = Some("Merge requires a single contiguous selection".to_string());
            cx.notify();
            return;
        }

        // Canonicalize selection range
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();

        // Guard: must select more than one cell
        if min_row == max_row && min_col == max_col {
            self.status_message = Some("Select a range of cells to merge".to_string());
            cx.notify();
            return;
        }

        let sheet = self.sheet(cx);

        // Overlap check: verify no partially-overlapping merges
        for merge in &sheet.merged_regions {
            let overlap_row = merge.start.0 <= max_row && merge.end.0 >= min_row;
            let overlap_col = merge.start.1 <= max_col && merge.end.1 >= min_col;
            if overlap_row && overlap_col {
                // Merge overlaps our selection - check if fully contained
                let fully_contained = merge.start.0 >= min_row
                    && merge.end.0 <= max_row
                    && merge.start.1 >= min_col
                    && merge.end.1 <= max_col;
                if !fully_contained {
                    self.status_message =
                        Some("Selection overlaps existing merged cells. Unmerge first.".to_string());
                    cx.notify();
                    return;
                }
            }
        }

        // Data-loss scan: check all cells except new origin
        let mut affected: Vec<String> = Vec::new();
        for r in min_row..=max_row {
            for c in min_col..=max_col {
                if (r, c) == (min_row, min_col) {
                    continue; // origin is kept
                }
                let raw = sheet.get_raw(r, c);
                if !raw.is_empty() {
                    affected.push(format!("{}{}", Self::col_letter(c), r + 1));
                }
            }
        }

        // Store range for both paths
        self.merge_confirm.merge_range = Some(((min_row, min_col), (max_row, max_col)));

        if affected.is_empty() {
            // No data loss — merge directly
            self.merge_cells_confirmed(cx);
        } else {
            // Show data-loss warning dialog
            self.merge_confirm.affected_cells = affected;
            self.merge_confirm.visible = true;
            cx.notify();
        }
    }

    /// Execute the merge after confirmation (or directly when no data loss).
    pub fn merge_cells_confirmed(&mut self, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::MergeCells);
        let ((min_row, min_col), (max_row, max_col)) = match self.merge_confirm.merge_range.take()
        {
            Some(range) => range,
            None => return,
        };

        let sheet_index = self.sheet_index(cx);

        // Snapshot before state
        let before = self.sheet(cx).merged_regions.clone();

        // Collect values to clear
        let mut cleared_values: Vec<(usize, usize, String)> = Vec::new();
        {
            let sheet = self.sheet(cx);
            for r in min_row..=max_row {
                for c in min_col..=max_col {
                    if (r, c) == (min_row, min_col) {
                        continue;
                    }
                    let raw = sheet.get_raw(r, c);
                    if !raw.is_empty() {
                        cleared_values.push((r, c, raw));
                    }
                }
            }
        }

        // Remove any existing merges fully inside the selection
        let contained_origins: Vec<(usize, usize)> = {
            let sheet = self.sheet(cx);
            sheet
                .merged_regions
                .iter()
                .filter(|m| {
                    m.start.0 >= min_row
                        && m.end.0 <= max_row
                        && m.start.1 >= min_col
                        && m.end.1 <= max_col
                })
                .map(|m| m.start)
                .collect()
        };
        for origin in contained_origins {
            self.active_sheet_mut(cx, |sheet| {
                sheet.remove_merge(origin);
            });
        }

        // Clear non-origin cell values
        self.wb_mut(cx, |wb| wb.begin_batch());
        for (row, col, _) in &cleared_values {
            self.set_cell_value(*row, *col, "", cx);
        }
        self.end_batch_and_broadcast(cx);

        // Add new merge
        self.active_sheet_mut(cx, |sheet| {
            let _ = sheet.add_merge(MergedRegion::new(min_row, min_col, max_row, max_col));
        });

        // Snapshot after state
        let after = self.sheet(cx).merged_regions.clone();

        // Build range ref for status message
        let range_ref = format!(
            "{}{}:{}{}",
            Self::col_letter(min_col),
            min_row + 1,
            Self::col_letter(max_col),
            max_row + 1,
        );

        // Record undo
        self.history.record_action_with_provenance(
            UndoAction::SetMerges {
                sheet_index,
                before,
                after,
                cleared_values,
                description: format!("Merge {}", range_ref),
            },
            None,
        );

        self.is_modified = true;
        self.status_message = Some(format!("Merged {}", range_ref));

        // Snap selection to merged range
        self.view_state.selected = (min_row, min_col);
        self.view_state.selection_end = Some((max_row, max_col));

        // Reset dialog state
        self.merge_confirm = Default::default();
        cx.notify();
    }

    /// Unmerge all merged regions that overlap the current selection.
    pub fn unmerge_cells(&mut self, cx: &mut Context<Self>) {
        self.set_repeat(RepeatAction::UnmergeCells);
        let sheet_index = self.sheet_index(cx);

        // Collect all merges that overlap any selection range
        let selection_ranges = self.all_selection_ranges();
        let mut origins_to_remove: Vec<(usize, usize)> = Vec::new();

        {
            let sheet = self.sheet(cx);
            for merge in &sheet.merged_regions {
                for &((min_row, min_col), (max_row, max_col)) in &selection_ranges {
                    let overlap_row = merge.start.0 <= max_row && merge.end.0 >= min_row;
                    let overlap_col = merge.start.1 <= max_col && merge.end.1 >= min_col;
                    if overlap_row && overlap_col {
                        if !origins_to_remove.contains(&merge.start) {
                            origins_to_remove.push(merge.start);
                        }
                    }
                }
            }
        }

        if origins_to_remove.is_empty() {
            self.status_message = Some("No merged cells in selection".to_string());
            cx.notify();
            return;
        }

        // Snapshot before state
        let before = self.sheet(cx).merged_regions.clone();

        // Remove all collected merges
        for origin in &origins_to_remove {
            self.active_sheet_mut(cx, |sheet| {
                sheet.remove_merge(*origin);
            });
        }

        // Snapshot after state
        let after = self.sheet(cx).merged_regions.clone();

        let count = origins_to_remove.len();

        // Record undo
        self.history.record_action_with_provenance(
            UndoAction::SetMerges {
                sheet_index,
                before,
                after,
                cleared_values: vec![], // unmerge doesn't clear values
                description: format!(
                    "Unmerge {} region{}",
                    count,
                    if count == 1 { "" } else { "s" }
                ),
            },
            None,
        );

        self.is_modified = true;
        self.status_message = Some(format!(
            "Unmerged {} region{}",
            count,
            if count == 1 { "" } else { "s" }
        ));
        cx.notify();
    }
}

impl Spreadsheet {
    /// Calculate which borders to draw for a selected cell.
    /// Returns (top, right, bottom, left) indicating which borders to draw.
    ///
    /// Strategy:
    /// - Always draw right+bottom (internal gridlines within selection)
    /// - Draw top only if cell above is NOT selected (outer edge)
    /// - Draw left only if cell to left is NOT selected (outer edge)
    /// This maintains the grid appearance while avoiding double borders at edges.
    pub fn selection_borders(&self, row: usize, col: usize) -> (bool, bool, bool, bool) {
        // Check if adjacent cells are also selected
        let cell_above_selected = row > 0 && self.is_selected(row - 1, col);
        let cell_left_selected = col > 0 && self.is_selected(row, col - 1);

        // Top/left: only at outer edges of selection
        let top = !cell_above_selected;
        let left = !cell_left_selected;

        // Right/bottom: always draw for internal gridlines
        let right = true;
        let bottom = true;

        (top, right, bottom, left)
    }
    /// Compute which user-defined borders to draw for a cell using adjacency logic.
    ///
    /// Returns (top, right, bottom, left) flags indicating which borders to draw.
    /// Uses the precedence rule: right/bottom takes precedence over left/top of adjacent cell.
    ///
    /// - Own right and bottom: always draw if set
    /// - Own top: only draw if cell above has no bottom border
    /// - Own left: only draw if cell to left has no right border
    pub fn cell_user_borders(
        &self, row: usize, col: usize, cx: &App,
        boundary_bottom: bool, boundary_right: bool,
    ) -> (CellBorder, CellBorder, CellBorder, CellBorder) {
        #[cfg(debug_assertions)]
        self.debug_border_call_count.set(self.debug_border_call_count.get() + 1);

        // Single-ownership model: each cell draws only its TOP and LEFT borders.
        // Right/bottom borders are drawn by the neighboring cell as their left/top.
        // Each edge is resolved as max(own_side, neighbor_opposite_side) so both
        // cells' border settings contribute, but only one cell draws the line.
        //
        // Exception: at the viewport boundary (last visible row/col), this cell
        // also draws BOTTOM (boundary_bottom) or RIGHT (boundary_right) because
        // the neighbor that would normally own that edge isn't rendered.
        //
        // For merged cells: interior cells draw nothing. Perimeter cells draw only
        // top (if on merge top edge) and left (if on merge left edge), resolved
        // with the neighboring cell/merge's opposing border. At viewport boundaries,
        // perimeter cells also draw bottom/right for merge edges that touch the boundary.

        let sheet = self.sheet(cx);

        // Helper: effective border contribution for a cell on a given side,
        // accounting for merges (interior cells contribute None, perimeter cells
        // contribute the merge's resolved edge border).
        let effective_side = |r: usize, c: usize, side: u8| -> CellBorder {
            // side: 0=top, 1=right, 2=bottom, 3=left
            if let Some(m) = sheet.get_merge(r, c) {
                let on_perimeter = match side {
                    0 => r == m.start.0,  // top
                    1 => c == m.end.1,    // right
                    2 => r == m.end.0,    // bottom
                    3 => c == m.start.1,  // left
                    _ => false,
                };
                if !on_perimeter {
                    return CellBorder::default(); // interior: no contribution
                }
                let (rt, rr, rb, rl) = sheet.resolve_merge_borders(m);
                match side {
                    0 => rt,
                    1 => rr,
                    2 => rb,
                    3 => rl,
                    _ => CellBorder::default(),
                }
            } else {
                let fmt = sheet.get_format(r, c);
                match side {
                    0 => fmt.border_top,
                    1 => fmt.border_right,
                    2 => fmt.border_bottom,
                    3 => fmt.border_left,
                    _ => CellBorder::default(),
                }
            }
        };

        // Check if this cell is a merge interior (not on any perimeter edge)
        let none = CellBorder::default();
        if let Some(m) = sheet.get_merge(row, col) {
            let on_edge = row == m.start.0 || row == m.end.0
                       || col == m.start.1 || col == m.end.1;
            if !on_edge {
                return (none, none, none, none); // interior: no borders
            }
        }

        // Resolve TOP edge: max(my_top, above_neighbor_bottom)
        let top = {
            let my_top = effective_side(row, col, 0);
            let above_bottom = if row > 0 {
                effective_side(row - 1, col, 2)
            } else {
                none
            };
            max_border(my_top, above_bottom)
        };

        // Resolve LEFT edge: max(my_left, left_neighbor_right)
        let left = {
            let my_left = effective_side(row, col, 3);
            let left_right = if col > 0 {
                effective_side(row, col - 1, 1)
            } else {
                none
            };
            max_border(my_left, left_right)
        };

        // Resolve BOTTOM edge: only at viewport boundary (last visible row)
        let bottom = if boundary_bottom {
            let my_bottom = effective_side(row, col, 2);
            let below_top = if row + 1 < NUM_ROWS {
                effective_side(row + 1, col, 0)
            } else {
                none
            };
            max_border(my_bottom, below_top)
        } else {
            none
        };

        // Resolve RIGHT edge: only at viewport boundary (last visible col)
        let right = if boundary_right {
            let my_right = effective_side(row, col, 1);
            let right_left = if col + 1 < NUM_COLS {
                effective_side(row, col + 1, 3)
            } else {
                none
            };
            max_border(my_right, right_left)
        } else {
            none
        };

        (top, right, bottom, left)
    }
    /// Check if any user-defined border is set for this cell
    pub fn has_user_borders(&self, row: usize, col: usize, cx: &App) -> bool {
        let format = self.sheet(cx).get_format(row, col);
        format.border_top.is_set() ||
        format.border_right.is_set() ||
        format.border_bottom.is_set() ||
        format.border_left.is_set()
    }
}

/// Decide which ranges a formatting command writes to.
///
/// Split out from `Spreadsheet::format_apply_ranges` so the clamping maths is
/// testable without a Window: everything here is arithmetic over the
/// selection and the populated bounding box.
///
/// `populated` is a closure because computing it walks the cell map, and the
/// common case (a normal-sized selection) never needs it.
fn plan_format_ranges(
    ranges: Vec<((usize, usize), (usize, usize))>,
    populated: impl FnOnce() -> Option<(usize, usize, usize, usize)>,
) -> Vec<((usize, usize), (usize, usize))> {
    let total: usize = ranges
        .iter()
        .map(|((r0, c0), (r1, c1))| {
            (r1.saturating_sub(*r0) + 1).saturating_mul(c1.saturating_sub(*c0) + 1)
        })
        .sum();

    if total <= MAX_FORMAT_APPLY_CELLS {
        return ranges;
    }

    let Some((ur0, uc0, ur1, uc1)) = populated() else {
        // Nothing on the sheet: formatting empty space would be invisible and
        // unbounded, so there is nothing worth writing.
        return Vec::new();
    };

    ranges
        .iter()
        .filter_map(|((r0, c0), (r1, c1))| {
            let a0 = (*r0).max(ur0);
            let b0 = (*c0).max(uc0);
            let a1 = (*r1).min(ur1);
            let b1 = (*c1).min(uc1);
            (a0 <= a1 && b0 <= b1).then_some(((a0, b0), (a1, b1)))
        })
        .collect()
}

#[cfg(test)]
mod format_apply_range_tests {
    // Explicit import: this module glob-imports gpui, whose `test` attribute
    // macro would otherwise shadow the built-in one.
    use super::{plan_format_ranges, MAX_FORMAT_APPLY_CELLS};

    const SHEET: (usize, usize) = (65_536, 256);

    fn select_all() -> Vec<((usize, usize), (usize, usize))> {
        vec![((0, 0), (SHEET.0 - 1, SHEET.1 - 1))]
    }

    /// The whole point: Ctrl+A then Ctrl+B must not touch 16.7M coordinates.
    #[test]
    fn select_all_clamps_to_populated_cells() {
        let planned = plan_format_ranges(select_all(), || Some((0, 0, 9, 3)));
        assert_eq!(planned, vec![((0, 0), (9, 3))]);
    }

    /// A selection a human could plausibly drag is applied exactly as given,
    /// so formatting an empty block then typing into it still works.
    #[test]
    fn ordinary_selections_are_untouched() {
        let ranges = vec![((0, 0), (99, 9)), ((200, 0), (204, 2))];
        let planned = plan_format_ranges(ranges.clone(), || {
            panic!("must not scan the cell map for a small selection")
        });
        assert_eq!(planned, ranges);
    }

    /// Exactly at the cap is still "ordinary" — the clamp is strictly above.
    #[test]
    fn cap_boundary_is_inclusive() {
        let rows = MAX_FORMAT_APPLY_CELLS / 10;
        let ranges = vec![((0, 0), (rows - 1, 9))];
        assert_eq!(
            plan_format_ranges(ranges.clone(), || panic!("should not scan at the cap")),
            ranges
        );
    }

    #[test]
    fn empty_sheet_writes_nothing() {
        assert!(plan_format_ranges(select_all(), || None).is_empty());
    }

    /// A discontiguous selection keeps only the parts that overlap content,
    /// and drops ranges that lie entirely outside it.
    #[test]
    fn disjoint_ranges_are_clipped_individually() {
        let ranges = vec![
            ((0, 0), (SHEET.0 - 1, 5)),   // overlaps content
            ((50_000, 100), (60_000, 120)), // far below/right of everything
        ];
        let planned = plan_format_ranges(ranges, || Some((2, 1, 40, 4)));
        assert_eq!(planned, vec![((2, 1), (40, 4))]);
    }

    /// Whole-column select (a real gesture: click the column header) is over
    /// the cap on its own and clamps to the rows that hold data.
    #[test]
    fn whole_column_clamps_to_its_data() {
        let planned = plan_format_ranges(vec![((0, 2), (SHEET.0 - 1, 2))], || Some((0, 0, 500, 8)));
        assert_eq!(planned, vec![((0, 2), (500, 2))]);
    }
}
