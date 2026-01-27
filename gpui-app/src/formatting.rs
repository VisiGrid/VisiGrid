//! Cell formatting methods for Spreadsheet
//!
//! This module contains all format setters (bold, italic, alignment, etc.)

use gpui::*;
use visigrid_engine::cell::{Alignment, CellBorder, NumberFormat, TextOverflow, VerticalAlignment};

use crate::app::{Spreadsheet, TriState, SelectionFormatState};
use crate::history::{CellFormatPatch, FormatActionKind};

/// Border application mode
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BorderApplyMode {
    /// Apply thin black borders to all 4 edges of each selected cell
    All,
    /// Apply thin black borders only to the outer perimeter of the selection
    Outline,
    /// Clear all borders from selected cells
    Clear,
}

impl Spreadsheet {
    /// Compute format state for the current selection (tri-state resolution)
    pub fn selection_format_state(&self, cx: &App) -> SelectionFormatState {
        let mut state = SelectionFormatState::default();
        let mut first = true;
        let mut last_display: Option<String> = None;

        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    state.cell_count += 1;
                    let raw = self.sheet(cx).get_raw(row, col);
                    let display = self.sheet(cx).get_display(row, col);
                    let format = self.sheet(cx).get_format(row, col);

                    if first {
                        state.raw_value = TriState::Uniform(raw.clone());
                        state.bold = TriState::Uniform(format.bold);
                        state.italic = TriState::Uniform(format.italic);
                        state.underline = TriState::Uniform(format.underline);
                        state.font_family = TriState::Uniform(format.font_family.clone());
                        state.alignment = TriState::Uniform(format.alignment);
                        state.vertical_alignment = TriState::Uniform(format.vertical_alignment);
                        state.text_overflow = TriState::Uniform(format.text_overflow);
                        state.number_format = TriState::Uniform(format.number_format);
                        last_display = Some(display);
                        first = false;
                    } else {
                        state.raw_value = state.raw_value.combine(&raw);
                        state.bold = state.bold.combine(&format.bold);
                        state.italic = state.italic.combine(&format.italic);
                        state.underline = state.underline.combine(&format.underline);
                        state.font_family = state.font_family.combine(&format.font_family);
                        state.alignment = state.alignment.combine(&format.alignment);
                        state.vertical_alignment = state.vertical_alignment.combine(&format.vertical_alignment);
                        state.text_overflow = state.text_overflow.combine(&format.text_overflow);
                        state.number_format = state.number_format.combine(&format.number_format);
                        last_display = Some(display);
                    }
                }
            }
        }

        // For single cell, show display value
        if matches!(state.raw_value, TriState::Uniform(_)) {
            state.display_value = last_display;
        }

        state
    }

    /// Set bold on all selected cells (explicit value, not toggle)
    pub fn set_bold(&mut self, value: bool, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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

    /// Set font family on all selected cells
    pub fn set_font_family_selection(&mut self, font: Option<String>, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
            };
            let desc = format!("Align {}", align_name);
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Alignment, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Set vertical alignment on all selected cells
    pub fn set_vertical_alignment_selection(&mut self, valign: VerticalAlignment, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    self.active_sheet_mut(cx, |s| s.set_number_format(row, col, format));
                    let after = self.sheet(cx).get_format(row, col);
                    if before != after {
                        patches.push(CellFormatPatch { row, col, before, after });
                    }
                }
            }
        }
        let count = patches.len();
        if count > 0 {
            let format_name = match format {
                NumberFormat::General => "General",
                NumberFormat::Number { .. } => "Number",
                NumberFormat::Currency { .. } => "Currency",
                NumberFormat::Percent { .. } => "Percent",
                NumberFormat::Date { .. } => "Date",
                NumberFormat::Time => "Time",
                NumberFormat::DateTime => "DateTime",
            };
            let desc = format!("{} format", format_name);
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::NumberFormat, desc.clone());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }

    /// Adjust decimal places on selected cells - uses DecimalPlaces kind for coalescing
    pub fn adjust_decimals_selection(&mut self, delta: i8, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    let new_format = match before.number_format {
                        NumberFormat::Number { decimals } => {
                            let new_dec = (decimals as i8 + delta).clamp(0, 10) as u8;
                            Some(NumberFormat::Number { decimals: new_dec })
                        }
                        NumberFormat::Currency { decimals } => {
                            let new_dec = (decimals as i8 + delta).clamp(0, 10) as u8;
                            Some(NumberFormat::Currency { decimals: new_dec })
                        }
                        NumberFormat::Percent { decimals } => {
                            let new_dec = (decimals as i8 + delta).clamp(0, 10) as u8;
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
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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

    /// Apply borders to all selected cells with canonicalization.
    ///
    /// Canonicalization: UI commands set BOTH sides of every shared edge they touch
    /// to prevent conflicting border states from normal use.
    pub fn apply_borders(&mut self, mode: BorderApplyMode, cx: &mut Context<Self>) {
        let thin = CellBorder::thin();
        let none = CellBorder::default();
        let mut patches = Vec::new();

        // For each selection range, apply borders with proper canonicalization
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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

        let count = patches.len();
        if count > 0 {
            let desc = match mode {
                BorderApplyMode::All => "All borders",
                BorderApplyMode::Outline => "Outline",
                BorderApplyMode::Clear => "Clear borders",
            };
            self.history.record_format(self.sheet_index(cx), patches, FormatActionKind::Border, desc.to_string());
            self.is_modified = true;
            self.status_message = Some(format!("{} → {} cell{}", desc, count, if count == 1 { "" } else { "s" }));
        }
        cx.notify();
    }
}
