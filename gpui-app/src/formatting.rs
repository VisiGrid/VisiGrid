//! Cell formatting methods for Spreadsheet
//!
//! This module contains all format setters (bold, italic, alignment, etc.)
//! and merge/unmerge cell operations.

use gpui::*;
use visigrid_engine::cell::{Alignment, CellBorder, CellFormat, NumberFormat, TextOverflow, VerticalAlignment};
use visigrid_engine::sheet::MergedRegion;

use crate::app::{Spreadsheet, TriState, SelectionFormatState};
use crate::history::{CellFormatPatch, FormatActionKind, UndoAction};
use crate::mode::Mode;

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
                        state.strikethrough = TriState::Uniform(format.strikethrough);
                        state.font_family = TriState::Uniform(format.font_family.clone());
                        state.alignment = TriState::Uniform(format.alignment);
                        state.vertical_alignment = TriState::Uniform(format.vertical_alignment);
                        state.text_overflow = TriState::Uniform(format.text_overflow);
                        state.number_format = TriState::Uniform(format.number_format);
                        state.background_color = TriState::Uniform(format.background_color);
                        state.font_size = TriState::Uniform(format.font_size);
                        state.font_color = TriState::Uniform(format.font_color);
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
                        last_display = Some(display);
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

    /// Set strikethrough on all selected cells (explicit value, not toggle)
    pub fn set_strikethrough(&mut self, value: bool, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
                Alignment::CenterAcrossSelection => "Center Across",
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
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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

    /// Set font size on all selected cells
    pub fn set_font_size_selection(&mut self, size: Option<f32>, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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

    /// Start Format Painter: capture the active cell's format.
    pub fn start_format_painter(&mut self, cx: &mut Context<Self>) {
        let (row, col) = self.view_state.selected;
        let format = self.sheet(cx).get_format(row, col);
        self.format_painter_format = Some(format);
        self.mode = crate::mode::Mode::FormatPainter;
        self.status_message = Some("Format Painter: click a cell to apply \u{00b7} Esc to cancel".to_string());
        cx.notify();
    }

    /// Apply Format Painter: set captured format on current selection.
    pub fn apply_format_painter(&mut self, cx: &mut Context<Self>) {
        let format = match self.format_painter_format.take() {
            Some(f) => f,
            None => return,
        };
        self.mode = crate::mode::Mode::Navigation;

        let mut patches = Vec::new();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    let before = self.sheet(cx).get_format(row, col);
                    if before != format {
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
        cx.notify();
    }

    /// Cancel Format Painter mode.
    pub fn cancel_format_painter(&mut self, cx: &mut Context<Self>) {
        self.format_painter_format = None;
        self.mode = crate::mode::Mode::Navigation;
        self.status_message = None;
        cx.notify();
    }

    /// Clear all formatting on selected cells, resetting to CellFormat::default().
    /// Records a single undo step regardless of cell count.
    pub fn clear_formatting_selection(&mut self, cx: &mut Context<Self>) {
        let mut patches = Vec::new();
        let default = CellFormat::default();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
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
