//! Clipboard operations for Spreadsheet.
//!
//! This module contains:
//! - InternalClipboard struct for tracking copied cell data
//! - Copy, cut, paste operations
//! - Paste Values (computed values only, no formulas)
//! - Delete selection

use gpui::*;
use visigrid_engine::cell::{Alignment, BorderStyle, CellBorder, CellFormat, CellStyle, VerticalAlignment};
use visigrid_engine::formula::eval::Value;
use visigrid_engine::provenance::{MutationOp, PasteMode, ClearMode};
use visigrid_engine::sheet::MergedRegion;

use visigrid_io::csv as csv_io;

use crate::app::{Spreadsheet, NUM_COLS, NUM_ROWS};
use crate::history::{CommentPatch, CellChange, CellFormatPatch, FormatActionKind, UndoAction};

/// Avoid accidental multi-gigabyte allocations when a whole row/column is selected.
const MAX_PICTURE_CELLS: usize = 10_000;
const MAX_PICTURE_DIMENSION: f32 = 8_192.0;
const MAX_PICTURE_PIXELS: f32 = 32_000_000.0;

pub(crate) fn svg_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

pub(crate) fn opaque_rgb(rgba: [u8; 4]) -> [u8; 3] {
    let alpha = rgba[3] as u16;
    let blend =
        |channel: u8| -> u8 { (((channel as u16 * alpha) + (255 * (255 - alpha))) / 255) as u8 };
    [blend(rgba[0]), blend(rgba[1]), blend(rgba[2])]
}

fn css_rgb(rgb: [u8; 3]) -> String {
    format!("#{:02x}{:02x}{:02x}", rgb[0], rgb[1], rgb[2])
}

fn semantic_fill(style: CellStyle) -> Option<[u8; 3]> {
    match style {
        CellStyle::None | CellStyle::Total => None,
        CellStyle::Error => Some([254, 226, 226]),
        CellStyle::Warning => Some([254, 243, 199]),
        CellStyle::Success => Some([220, 252, 231]),
        CellStyle::Input => Some([219, 234, 254]),
        CellStyle::Note => Some([249, 250, 251]),
    }
}

fn border_width(style: BorderStyle) -> f32 {
    match style {
        BorderStyle::None => 0.0,
        BorderStyle::Thin => 1.0,
        BorderStyle::Medium => 2.0,
        BorderStyle::Thick => 3.0,
    }
}

fn push_svg_border(svg: &mut String, border: CellBorder, x1: f32, y1: f32, x2: f32, y2: f32) {
    let width = border_width(border.style);
    if width == 0.0 {
        return;
    }
    let color = css_rgb(border.color.map(opaque_rgb).unwrap_or([0, 0, 0]));
    use std::fmt::Write as _;
    let _ = write!(
        svg,
        r#"<line x1="{x1:.2}" y1="{y1:.2}" x2="{x2:.2}" y2="{y2:.2}" stroke="{color}" stroke-width="{width:.2}"/>"#,
    );
}

/// Internal clipboard for tracking copied cell data.
/// Stores both raw formulas (for normal paste) and typed values (for paste values).
#[derive(Debug, Clone)]
pub struct InternalClipboard {
    /// Tab-separated raw values (formulas/text) for normal paste + system clipboard
    pub raw_tsv: String,
    /// Exact cell boundaries; text may itself contain tabs or newlines.
    pub raw_cells: Vec<Vec<String>>,
    /// Typed computed values for Paste Values (2D grid aligned to copied rectangle)
    pub values: Vec<Vec<Value>>,
    /// Cell formats for Paste Formats (2D grid with same dimensions as values)
    /// Every position gets a CellFormat, even if default (rectangular, not sparse).
    pub formats: Vec<Vec<CellFormat>>,
    pub comments: Vec<Vec<Option<visigrid_engine::cell::CellComment>>>,
    /// Top-left cell position of the copied region (for reference adjustment)
    pub source: (usize, usize),
    /// Canonical source record for each copied row, captured before any view changes.
    pub source_rows: Vec<usize>,
    pub source_formulas: Vec<Vec<bool>>,
    /// Unique ID written to clipboard metadata for reliable internal detection.
    /// On paste, we check if clipboard metadata contains this ID to distinguish
    /// internal copies from external clipboard content (even if text matches).
    pub id: u128,
    /// Merged regions from the copied area, stored with coordinates
    /// relative to the clipboard's top-left (0,0).
    /// Empty when copied from a filtered view.
    pub merges: Vec<MergedRegion>,
    /// When this clipboard entry was created (for time-bounded Wayland fallback)
    pub created_at: std::time::Instant,
}

fn full_paste_lines(text: &str, internal: bool) -> Vec<&str> {
    if internal { text.split('\n').collect() } else { text.lines().collect() }
}

#[cfg(test)]
mod comment_clipboard_tests {
    use super::full_paste_lines;
    #[test]
    fn comment_only_cells_keep_their_clipboard_rows() {
        assert_eq!(full_paste_lines("", true), vec![""]);
        assert_eq!(full_paste_lines("value\n\n", true), vec!["value", "", ""]);
        assert_eq!(full_paste_lines("value\r\n", false), vec!["value"]);
    }
}

impl Spreadsheet {
    // Clipboard
    /// Copy the selected range as a portable PNG image.
    ///
    /// The picture is rebuilt from the sheet model instead of cropping the
    /// viewport, so off-screen cells and the current selection highlight are
    /// handled correctly.
    pub fn copy_as_picture(&mut self, cx: &mut Context<Self>) {
        use std::fmt::Write as _;

        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let row_count = max_row.saturating_sub(min_row) + 1;
        let col_count = max_col.saturating_sub(min_col) + 1;
        if row_count.saturating_mul(col_count) > MAX_PICTURE_CELLS {
            self.status_message = Some(format!(
                "Selection is too large to copy as a picture (maximum {MAX_PICTURE_CELLS} cells)"
            ));
            cx.notify();
            return;
        }

        let rows: Vec<(usize, usize, f32)> = (min_row..=max_row)
            .filter(|view_row| self.row_view.is_view_row_visible(*view_row))
            .map(|view_row| {
                let data_row = self.row_view.view_to_data(view_row);
                (view_row, data_row, self.row_height(view_row).max(1.0))
            })
            .collect();
        let cols: Vec<(usize, f32)> = (min_col..=max_col)
            .map(|col| (col, self.col_width(col).max(1.0)))
            .collect();

        let width: f32 = cols.iter().map(|(_, width)| width).sum();
        let height: f32 = rows.iter().map(|(_, _, height)| height).sum();
        if rows.is_empty() || cols.is_empty() || width <= 0.0 || height <= 0.0 {
            self.status_message = Some("Nothing to copy as a picture".to_string());
            cx.notify();
            return;
        }

        let show_gridlines = match &crate::settings::user_settings(cx).appearance.show_gridlines {
            crate::settings::Setting::Value(value) => *value,
            crate::settings::Setting::Inherit => true,
        };
        let mut svg = format!(
            r##"<svg xmlns="http://www.w3.org/2000/svg" width="{width:.2}" height="{height:.2}" viewBox="0 0 {width:.2} {height:.2}"><rect width="100%" height="100%" fill="#ffffff"/><g shape-rendering="crispEdges">"##,
        );

        let mut y = 0.0;
        for (row_index, (_, data_row, cell_height)) in rows.iter().enumerate() {
            let mut x = 0.0;
            for (col_index, (col, cell_width)) in cols.iter().enumerate() {
                let sheet = self.sheet(cx);
                let merge = sheet.get_merge(*data_row, *col).cloned();
                let (format_row, format_col) = merge
                    .as_ref()
                    .map(|region| region.start)
                    .unwrap_or((*data_row, *col));
                let format = self.effective_format_cached(format_row, format_col, cx);
                let fill = format
                    .background_color
                    .map(opaque_rgb)
                    .or_else(|| semantic_fill(format.cell_style))
                    .unwrap_or([255, 255, 255]);
                let _ = write!(
                    svg,
                    r##"<rect x="{x:.2}" y="{y:.2}" width="{cell_width:.2}" height="{cell_height:.2}" fill="{}"/>"##,
                    css_rgb(fill),
                );

                let same_merge_right = merge.as_ref().is_some_and(|region| *col < region.end.1);
                let same_merge_bottom = merge
                    .as_ref()
                    .is_some_and(|region| *data_row < region.end.0);
                if show_gridlines {
                    let grid = "#d9d9d9";
                    if !same_merge_right {
                        let _ = write!(
                            svg,
                            r#"<line x1="{:.2}" y1="{y:.2}" x2="{:.2}" y2="{:.2}" stroke="{grid}" stroke-width="1"/>"#,
                            x + cell_width,
                            x + cell_width,
                            y + cell_height
                        );
                    }
                    if !same_merge_bottom {
                        let _ = write!(
                            svg,
                            r#"<line x1="{x:.2}" y1="{:.2}" x2="{:.2}" y2="{:.2}" stroke="{grid}" stroke-width="1"/>"#,
                            y + cell_height,
                            x + cell_width,
                            y + cell_height
                        );
                    }
                    if row_index == 0 {
                        let _ = write!(
                            svg,
                            r#"<line x1="{x:.2}" y1="{y:.2}" x2="{:.2}" y2="{y:.2}" stroke="{grid}" stroke-width="1"/>"#,
                            x + cell_width
                        );
                    }
                    if col_index == 0 {
                        let _ = write!(
                            svg,
                            r#"<line x1="{x:.2}" y1="{y:.2}" x2="{x:.2}" y2="{:.2}" stroke="{grid}" stroke-width="1"/>"#,
                            y + cell_height
                        );
                    }
                }

                push_svg_border(&mut svg, format.border_top, x, y, x + cell_width, y);
                push_svg_border(
                    &mut svg,
                    format.border_right,
                    x + cell_width,
                    y,
                    x + cell_width,
                    y + cell_height,
                );
                push_svg_border(
                    &mut svg,
                    format.border_bottom,
                    x,
                    y + cell_height,
                    x + cell_width,
                    y + cell_height,
                );
                push_svg_border(&mut svg, format.border_left, x, y, x, y + cell_height);

                let is_merge_hidden = sheet.is_merge_hidden(*data_row, *col);
                if !is_merge_hidden {
                    let mut display = if self.show_formulas() {
                        sheet.get_raw(*data_row, *col)
                    } else {
                        sheet.get_formatted_display(*data_row, *col)
                    };
                    if !self.show_zeros() && display == "0" {
                        display.clear();
                    }
                    if !display.is_empty() {
                        let text_width = merge.as_ref().map_or(*cell_width, |region| {
                            cols.iter()
                                .filter(|(candidate, _)| {
                                    *candidate >= region.start.1 && *candidate <= region.end.1
                                })
                                .map(|(_, width)| *width)
                                .sum::<f32>()
                                .max(*cell_width)
                        });
                        let text_height = merge.as_ref().map_or(*cell_height, |region| {
                            rows.iter()
                                .filter(|(_, candidate, _)| {
                                    *candidate >= region.start.0 && *candidate <= region.end.0
                                })
                                .map(|(_, _, height)| *height)
                                .sum::<f32>()
                                .max(*cell_height)
                        });
                        let computed = sheet.get_computed_value(*data_row, *col);
                        let alignment = match format.alignment {
                            Alignment::General if matches!(computed, Value::Number(_)) => {
                                Alignment::Right
                            }
                            Alignment::General => Alignment::Left,
                            other => other,
                        };
                        let (text_x, anchor) = match alignment {
                            Alignment::Right => (text_width - 4.0, "end"),
                            Alignment::Center | Alignment::CenterAcrossSelection => {
                                (text_width / 2.0, "middle")
                            }
                            Alignment::Left | Alignment::General => (4.0, "start"),
                        };
                        let font_size = self.cell_font.pixels(format.font_size, 1.0);
                        let (text_y, baseline) = match format.vertical_alignment {
                            VerticalAlignment::Top => (3.0, "hanging"),
                            VerticalAlignment::Middle => (text_height / 2.0, "central"),
                            VerticalAlignment::Bottom => (text_height - 3.0, "auto"),
                        };
                        let color =
                            css_rgb(format.font_color.map(opaque_rgb).unwrap_or([32, 32, 32]));
                        let family = svg_escape(self.cell_font_family(format.font_family.as_deref()).as_ref());
                        let weight = if format.bold || format.cell_style == CellStyle::Total {
                            "700"
                        } else {
                            "400"
                        };
                        let italic = if format.italic { "italic" } else { "normal" };
                        let decoration = match (format.underline, format.strikethrough) {
                            (true, true) => "underline line-through",
                            (true, false) => "underline",
                            (false, true) => "line-through",
                            (false, false) => "none",
                        };
                        let text = svg_escape(&display.replace(['\n', '\r'], " "));
                        let _ = write!(
                            svg,
                            r#"<svg x="{x:.2}" y="{y:.2}" width="{text_width:.2}" height="{text_height:.2}" overflow="hidden"><text x="{text_x:.2}" y="{text_y:.2}" text-anchor="{anchor}" dominant-baseline="{baseline}" font-family="{family}" font-size="{font_size:.2}" font-weight="{weight}" font-style="{italic}" text-decoration="{decoration}" fill="{color}">{text}</text></svg>"#,
                        );
                    }
                }

                x += cell_width;
            }
            y += cell_height;
        }
        svg.push_str("</g></svg>");

        let mut options = resvg::usvg::Options::default();
        options.fontdb_mut().load_system_fonts();
        // Pictures must use the same bundled fonts as the grid, even when they
        // aren't installed on the host OS.
        for font in crate::embedded_fonts() {
            options.fontdb_mut().load_font_data(font.into_owned());
        }
        let tree = match resvg::usvg::Tree::from_str(&svg, &options) {
            Ok(tree) => tree,
            Err(error) => {
                self.status_message = Some(format!("Could not render selection picture: {error}"));
                cx.notify();
                return;
            }
        };
        let scale = (MAX_PICTURE_DIMENSION / width)
            .min(MAX_PICTURE_DIMENSION / height)
            .min((MAX_PICTURE_PIXELS / (width * height)).sqrt())
            .min(1.0);
        let pixel_width = (width * scale).ceil().max(1.0) as u32;
        let pixel_height = (height * scale).ceil().max(1.0) as u32;
        let Some(mut pixmap) = resvg::tiny_skia::Pixmap::new(pixel_width, pixel_height) else {
            self.status_message = Some("Could not allocate selection picture".to_string());
            cx.notify();
            return;
        };
        resvg::render(
            &tree,
            resvg::tiny_skia::Transform::from_scale(scale, scale),
            &mut pixmap.as_mut(),
        );
        match pixmap.encode_png() {
            Ok(bytes) => {
                let image = gpui::Image::from_bytes(gpui::ImageFormat::Png, bytes);
                cx.write_to_clipboard(ClipboardItem::new_image(&image));
                self.status_message = Some(format!(
                    "Copied {} x {} selection as picture",
                    rows.len(),
                    cols.len()
                ));
            }
            Err(error) => {
                self.status_message = Some(format!("Could not encode selection picture: {error}"));
            }
        }
        cx.notify();
    }

    pub fn copy(&mut self, cx: &mut Context<Self>) {
        // If editing, copy selected text (or all if no selection)
        // This is text-only copy, not cell copy - no internal clipboard needed
        if self.mode.is_editing() {
            let text = if let Some((start_byte, end_byte)) = self.edit_selection_range() {
                // Byte-indexed selection
                let start = start_byte.min(self.edit_value.len());
                let end = end_byte.min(self.edit_value.len());
                self.edit_value[start..end].to_string()
            } else {
                self.edit_value.clone()
            };
            self.internal_clipboard = None;  // Text copy, not cell copy
            cx.write_to_clipboard(ClipboardItem::new_string(text));
            self.status_message = Some("Copied to clipboard".to_string());
            cx.notify();
            return;
        }

        if self.review_mode.is_some() {
            let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
            let mut tsv = String::new();
            for view_row in min_row..=max_row {
                if view_row > min_row {
                    tsv.push('\n');
                }
                let data_row = self.view_to_data(view_row, cx);
                for col in min_col..=max_col {
                    if col > min_col {
                        tsv.push('\t');
                    }
                    let sheet_id = self.sheet(cx).id;
                    let displayed = self
                        .review_endpoint_sheet_row(sheet_id, data_row)
                        .map(|(sheet, endpoint_row)| {
                            if self.show_formulas() {
                                sheet.get_raw(endpoint_row, col)
                            } else {
                                sheet.get_formatted_display(endpoint_row, col)
                            }
                        })
                        .unwrap_or_else(|| {
                            if self.show_formulas() {
                                self.sheet(cx).get_raw(data_row, col)
                            } else {
                                self.sheet(cx).get_formatted_display(data_row, col)
                            }
                        });
                    tsv.push_str(&displayed);
                }
            }
            self.internal_clipboard = None;
            cx.write_to_clipboard(ClipboardItem::new_string(tsv));
            self.status_message = Some("Copied displayed Review Mode values".to_string());
            cx.notify();
            return;
        }

        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let is_filtered = self.row_view.is_filtered();

        // Normalize selection to include full merge regions (only when not filtered)
        let (min_row, min_col, max_row, max_col) = if !is_filtered {
            let sheet = self.sheet(cx);
            let mut nr = min_row;
            let mut nc = min_col;
            let mut xr = max_row;
            let mut xc = max_col;
            // Expand to include any intersecting merges (one pass suffices since merges don't overlap)
            for merge in &sheet.merged_regions {
                let intersects = merge.end.0 >= nr && merge.start.0 <= xr
                              && merge.end.1 >= nc && merge.start.1 <= xc;
                if intersects {
                    nr = nr.min(merge.start.0);
                    nc = nc.min(merge.start.1);
                    xr = xr.max(merge.end.0);
                    xc = xc.max(merge.end.1);
                }
            }
            (nr, nc, xr, xc)
        } else {
            (min_row, min_col, max_row, max_col)
        };

        // Build tab-separated raw values (formulas) for system clipboard and normal paste
        // When filtered, only include visible rows
        let mut raw_tsv = String::new();
        let mut raw_cells = Vec::new();
        let mut values = Vec::new();
        let mut formats = Vec::new();
        let mut comments = Vec::new();
        let mut source_rows = Vec::new();
        let mut source_formulas = Vec::new();
        let mut first_row = true;
        let mut source_row = min_row; // Track first visible row for source

        for view_row in min_row..=max_row {
            // Skip hidden rows when filtered
            if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                continue;
            }

            // Convert view row to data row for sheet access
            let data_row = self.row_view.view_to_data(view_row);

            if first_row {
                source_row = data_row;
                first_row = false;
            } else {
                raw_tsv.push('\n');
            }

            let mut row_raw = Vec::new();
            let mut row_values = Vec::new();
            let mut row_formats = Vec::new();
            let mut row_comments = Vec::new();
            for col in min_col..=max_col {
                if col > min_col {
                    raw_tsv.push('\t');
                }
                let raw = self.sheet(cx).get_raw(data_row, col);
                raw_tsv.push_str(&raw);
                row_raw.push(raw);
                row_values.push(self.sheet(cx).get_computed_value(data_row, col));
                // Capture format for every cell position (rectangular, not sparse)
                row_formats.push(self.sheet(cx).get_format(data_row, col).clone());
                row_comments.push(self.sheet(cx).comment(data_row, col).cloned());
            }
            source_formulas.push((min_col..=max_col).map(|c| self.sheet(cx).get_cell_opt(data_row,c).is_some_and(|cell| cell.value().is_formula())).collect());
            source_rows.push(data_row);
            raw_cells.push(row_raw);
            values.push(row_values);
            formats.push(row_formats);
            comments.push(row_comments);
        }

        // Capture merge metadata (only when not filtered)
        let merges = if !is_filtered {
            let sheet = self.sheet(cx);
            let mut relative_merges = Vec::new();
            for merge in &sheet.merged_regions {
                // After normalization, all intersecting merges are fully contained
                let contained = merge.start.0 >= min_row && merge.end.0 <= max_row
                              && merge.start.1 >= min_col && merge.end.1 <= max_col;
                if contained {
                    relative_merges.push(MergedRegion::new(
                        merge.start.0 - min_row,
                        merge.start.1 - min_col,
                        merge.end.0 - min_row,
                        merge.end.1 - min_col,
                    ));
                }
            }
            relative_merges
        } else {
            Vec::new()
        };

        // Generate unique nonce for clipboard matching
        let id: u128 = rand::random();

        // A cell copy supersedes a format copied with Ctrl+Shift+C, so
        // Ctrl+Shift+V goes back to pasting values.
        if self.mode != crate::mode::Mode::FormatPainter {
            self.format_painter = None;
        }

        self.internal_clipboard = Some(InternalClipboard {
            raw_tsv: raw_tsv.clone(),
            raw_cells,
            values,
            formats,
            comments,
            source: (source_row, min_col),
            source_rows,
            source_formulas,
            id,
            merges,
            created_at: std::time::Instant::now(),
        });
        // Write clipboard with metadata ID for reliable internal detection
        let id_json = format!("\"{}\"", id);
        // A newline is a valid TSV record for one blank cell. Unlike a zero-byte
        // payload, it survives clipboard providers that treat empty text as absent.
        let clipboard_text = if raw_tsv.is_empty() { "\n".to_owned() } else { raw_tsv };
        cx.write_to_clipboard(ClipboardItem::new_string_with_json_metadata(clipboard_text, id_json));

        // Set visual range for dashed border overlay
        self.clipboard_visual_range = Some((min_row, min_col, max_row, max_col));
        self.clipboard_visual_sheet = Some(self.sheet(cx).id);
        self.start_marching_ants(cx);

        if is_filtered {
            self.status_message = Some("Copied visible rows to clipboard".to_string());
        } else {
            self.status_message = Some("Copied to clipboard".to_string());
        }
        cx.notify();
    }

    pub fn cut(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        if self.mode.is_editing() {
            self.copy(cx);
            if self.edit_selection_range().is_none() {
                self.edit_selection_anchor = Some(0);
                self.edit_cursor = self.edit_value.len();
            }
            self.backspace(cx);
            return;
        }
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.cut_table_view(cx);
            return;
        }
        // Block during preview mode
        if self.block_if_previewing(cx) { return; }
        if self.block_if_selection_in_pivot("cut", cx) { return; }
        let ((tr0,tc0),(tr1,tc1))=self.selection_range();
        if self.block_if_table_header(tr0,tc0,tr1,tc1,"cut",cx) { return; }

        self.copy(cx);

        // Clear the selected cells and record history (visible rows only when filtered)
        let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
        let is_filtered = self.row_view.is_filtered();

        // Normalize selection to include full merge regions (same expansion as copy)
        let (min_row, min_col, max_row, max_col) = if !is_filtered {
            let sheet = self.sheet(cx);
            let mut nr = min_row;
            let mut nc = min_col;
            let mut xr = max_row;
            let mut xc = max_col;
            for merge in &sheet.merged_regions {
                let intersects = merge.end.0 >= nr && merge.start.0 <= xr
                              && merge.end.1 >= nc && merge.start.1 <= xc;
                if intersects {
                    nr = nr.min(merge.start.0);
                    nc = nc.min(merge.start.1);
                    xr = xr.max(merge.end.0);
                    xc = xc.max(merge.end.1);
                }
            }
            (nr, nc, xr, xc)
        } else {
            (min_row, min_col, max_row, max_col)
        };

        let mut changes = Vec::new();
        let mut comment_patches = Vec::new();

        self.wb_mut(cx, |wb| wb.begin_batch());
        for view_row in min_row..=max_row {
            // Skip hidden rows when filtered
            if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                continue;
            }

            // Convert view row to data row for sheet access
            let data_row = self.row_view.view_to_data(view_row);

            for col in min_col..=max_col {
                let old_value = self.sheet(cx).get_raw(data_row, col);
                if !old_value.is_empty() {
                    changes.push(CellChange {
                        row: data_row, col, old_value, new_value: String::new(),
                    });
                }
                if let Some(before) = self.sheet(cx).comment(data_row, col).cloned() {
                    comment_patches.push(CommentPatch { remove_cell_on_undo: false, row: data_row, col, before: Some(before), after: None });
                    self.active_sheet_mut(cx, |s| s.set_comment(data_row, col, None));
                }
                self.set_cell_value(data_row, col, "", cx);
            }
        }
        self.end_batch_and_broadcast(cx);

        // Remove merges fully within the cut selection (move semantics, not filtered)
        let mut removed_any = false;
        let mut merges_before = Vec::new();
        let mut merges_after = Vec::new();

        if !is_filtered {
            merges_before = self.sheet(cx).merged_regions.clone();
            let origins_to_remove: Vec<(usize, usize)> = {
                let sheet = self.sheet(cx);
                sheet.merged_regions.iter()
                    .filter(|merge| {
                        merge.start.0 >= min_row && merge.end.0 <= max_row
                        && merge.start.1 >= min_col && merge.end.1 <= max_col
                    })
                    .map(|merge| merge.start)
                    .collect()
            };
            removed_any = !origins_to_remove.is_empty();
            for origin in origins_to_remove {
                self.active_sheet_mut(cx, |s| { let _ = s.remove_merge(origin); });
            }
            merges_after = self.sheet(cx).merged_regions.clone();
        }

        let sheet_index = self.sheet_index(cx);
        let mut actions = vec![UndoAction::Values { sheet_index, changes }];
        if removed_any {
            actions.push(UndoAction::SetMerges { sheet_index, before: merges_before, after: merges_after,
                cleared_values: vec![], description: "Cut: remove source merges".into() });
        }
        if !comment_patches.is_empty() {
            self.wb_mut(cx, |wb| wb.bump_revision_for_structure());
            actions.push(UndoAction::Comments { sheet_index, patches: comment_patches, description: "Cut comments".into() });
        }
        self.record_action_with_provenance(cx, UndoAction::Group { actions, description: "Cut".into() }, None);

        self.bump_cells_rev();  // Invalidate cell search cache
        self.is_modified = true;

        if is_filtered {
            // Check if there were merges in the cut range that we didn't remove
            let has_merges_in_range = self.sheet(cx).merged_regions.iter().any(|m| {
                m.end.0 >= min_row && m.start.0 <= max_row
                && m.end.1 >= min_col && m.start.1 <= max_col
            });
            if has_merges_in_range {
                self.status_message = Some("Cut visible rows (merged regions not moved in filtered view)".to_string());
            } else {
                self.status_message = Some("Cut visible rows to clipboard".to_string());
            }
        } else {
            self.status_message = Some("Cut to clipboard".to_string());
        }
        cx.notify();
    }

    pub fn paste(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if self.block_if_previewing_only(cx) { return; }

        // If editing, paste into the edit buffer instead
        if self.mode.is_editing() {
            self.paste_into_edit(cx);
            return;
        }

        // Purist setting: Ctrl+V pastes values only. Full paste remains
        // available via Paste Special > All (Ctrl+Alt+V).
        if crate::settings::user_settings(cx)
            .editing
            .paste_values_by_default
            .resolve(false)
        {
            self.paste_values(cx);
            return;
        }

        // Ctrl+V brings contents (formulas, comments, merges). Cells that are
        // already formatted keep their look; blank, unformatted cells take the
        // copied formatting, so pasting into empty space still looks right.
        self.paste_contents(false, cx);
    }

    /// Paste Special > All: contents plus the copied cells' formatting.
    /// Explicit, so it never inherits the values-only preference.
    pub fn paste_all(&mut self, cx: &mut Context<Self>) {
        self.paste_contents(true, cx);
    }

    /// Set a pasted cell's format, recording the change for undo. Unless
    /// `overwrite`, only a cell with no formatting of its own takes it.
    fn set_pasted_format(&mut self, row: usize, col: usize, format: CellFormat, overwrite: bool, patches: &mut Vec<CellFormatPatch>, cx: &mut Context<Self>) {
        let before = self.sheet(cx).get_format(row, col).clone();
        if !overwrite && before != CellFormat::default() {
            return;
        }
        if before != format {
            self.active_sheet_mut(cx, |s| s.set_format(row, col, format.clone()));
            patches.push(CellFormatPatch { remove_cell_on_undo: false, row, col, before, after: format });
        }
    }

    /// Full paste. Formatting comes only from an internal clipboard: with
    /// `all_formats` it replaces every destination format (Paste Special > All);
    /// without, it fills only unformatted cells (Ctrl+V).
    fn paste_contents(&mut self, all_formats: bool, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        let kind = if all_formats { TablePasteKind::All } else { TablePasteKind::Contents };
        if self.paste_table_headers(kind, cx) { return; }
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.paste_table_view(kind, cx);
            return;
        }
        if self.block_if_previewing(cx) { return; }
        if self.mode.is_editing() { self.paste_into_edit(cx); return; }
        // Read clipboard item to get both text and metadata
        let clipboard_item = cx.read_from_clipboard();
        let system_text = clipboard_item.as_ref().and_then(|item| item.text().map(|s| s.to_string()));
        let metadata = clipboard_item.as_ref().and_then(|item| item.metadata().cloned());

        // Determine if this is an internal paste (with formula adjustment) or external
        let is_internal = Self::is_internal_paste(
            self.internal_clipboard.as_ref(),
            system_text.as_deref(),
            metadata.as_deref(),
        );

        #[cfg(debug_assertions)]
        {
            let text_match = system_text.as_deref().map_or(false, |st| {
                self.internal_clipboard.as_ref().map_or(false, |ic| {
                    Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv)
                })
            });
            eprintln!("[paste] is_internal={}, metadata={:?}, text_match={}", is_internal, metadata.is_some(), text_match);
        }

        // Retain the exact internal rectangle, including blank trailing rows.
        let text = if is_internal {
            self.internal_clipboard.as_ref().map(|ic| ic.raw_tsv.clone())
        } else {
            system_text
        };

        if let Some(text) = text {
            let (start_row, start_col) = self.view_state.selected;
            let is_filtered = self.row_view.is_filtered();
            let mut changes = Vec::new();
            let mut comment_patches = Vec::new();
            let mut format_patches = Vec::new();
            let with_formats = is_internal;

            // For external pastes without tabs, try CSV-aware parsing (handles commas,
            // semicolons, pipes, and quoted fields). Only use the result if it found
            // multiple columns — otherwise fall through to existing tab-split behavior.
            //
            // Guard: if the sniffed delimiter is comma, check whether every line is
            // actually a formatted number (e.g. "6,601.43"). Commas inside numbers
            // are thousands separators, not field delimiters.
            let parsed_grid: Option<Vec<Vec<String>>> = if is_internal {
                self.internal_clipboard.as_ref().map(|ic| ic.raw_cells.clone())
            } else if !text.contains('\t') {
                let sniffed = csv_io::sniff_delimiter(&text);
                let grid = csv_io::parse_delimited_text(&text);
                let has_multi_col = grid.iter().any(|row| row.len() > 1);

                if has_multi_col && sniffed == b',' {
                    // If every line parses as a formatted number, the commas are
                    // thousands separators — don't split into columns.
                    let all_numbers = text.lines().all(|line| {
                        visigrid_engine::cell::try_parse_number(line.trim()).is_some()
                    });
                    if all_numbers { None } else { Some(grid) }
                } else if has_multi_col {
                    Some(grid)
                } else {
                    None
                }
            } else {
                None
            };

            // Check if clipboard is a single cell (1 line, no tabs, no CSV multi-col)
            // Internal TSV includes its full rectangle, even blank trailing rows
            // and a single empty cell carrying only a comment.
            let lines: Vec<&str> = full_paste_lines(&text, is_internal);
            let is_single_cell = parsed_grid.as_ref().map_or_else(
                || lines.len() == 1 && !lines[0].contains('\t'),
                |grid| grid.len() == 1 && grid[0].len() == 1,
            );

            // If single cell and multi-selection, broadcast to all selected cells
            if is_single_cell && self.is_multi_selection() {
                if self.block_if_selection_in_pivot("paste", cx) { return; }
                if self.block_selection_table_headers("paste", cx) { return; }
                let single_value = parsed_grid.as_ref().map_or_else(
                    || lines[0].to_string(), |grid| grid[0][0].clone(),
                );
                let primary_cell = self.view_state.selected;
                let primary_data_row = self.row_view.view_to_data(primary_cell.0);

                // Source cell position for formula rebasing (delta = target - source)
                let (src_data_row, src_col) = if is_internal {
                    if let Some(ic) = &self.internal_clipboard {
                        (ic.source.0 as i32, ic.source.1 as i32)
                    } else {
                        (primary_data_row as i32, primary_cell.1 as i32)
                    }
                } else {
                    (primary_data_row as i32, primary_cell.1 as i32)
                };

                // Collect all target cells (view_row, col) -> (data_row, col)
                let mut target_cells: Vec<(usize, usize)> = Vec::new();

                // Primary selection rectangle (filter to visible rows)
                let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
                for view_row in min_row..=max_row {
                    if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                        continue;
                    }
                    let data_row = self.row_view.view_to_data(view_row);
                    for col in min_col..=max_col {
                        target_cells.push((data_row, col));
                    }
                }

                // Additional selections (Ctrl+Click) - filter to visible rows
                for (sel_start, sel_end) in &self.view_state.additional_selections {
                    let end = sel_end.unwrap_or(*sel_start);
                    let min_r = sel_start.0.min(end.0);
                    let max_r = sel_start.0.max(end.0);
                    let min_c = sel_start.1.min(end.1);
                    let max_c = sel_start.1.max(end.1);
                    for view_row in min_r..=max_r {
                        if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                            continue;
                        }
                        let data_row = self.row_view.view_to_data(view_row);
                        for col in min_c..=max_c {
                            if !target_cells.contains(&(data_row, col)) {
                                target_cells.push((data_row, col));
                            }
                        }
                    }
                }

                let is_formula = single_value.starts_with('=');
                let mut values_grid: Vec<Vec<String>> = Vec::new();

                self.wb_mut(cx, |wb| wb.begin_batch());
                for (data_row, col) in &target_cells {
                    let old_value = self.sheet(cx).get_raw(*data_row, *col);

                    // For formulas, shift relative references based on delta from source cell
                    let new_value = if is_formula && is_internal {
                        let delta_row = *data_row as i32 - src_data_row;
                        let delta_col = *col as i32 - src_col;
                        self.adjust_formula_refs(&single_value, delta_row, delta_col)
                    } else {
                        single_value.clone()
                    };

                    if old_value != new_value {
                        changes.push(CellChange {
                            row: *data_row, col: *col, old_value, new_value: new_value.clone(),
                        });
                    }
                    self.set_cell_value(*data_row, *col, &new_value, cx);
                    if is_internal {
                        let after = self.internal_clipboard.as_ref().and_then(|ic| ic.comments.first()).and_then(|r| r.first()).cloned().flatten();
                        let before = self.sheet(cx).comment(*data_row, *col).cloned();
                        if before != after {
                            self.active_sheet_mut(cx, |s| s.set_comment(*data_row, *col, after.clone()));
                            comment_patches.push(CommentPatch { remove_cell_on_undo: false, row: *data_row, col: *col, before, after });
                        }
                    }
                    if with_formats {
                        if let Some(format) = self.internal_clipboard.as_ref().and_then(|ic| ic.formats.first()).and_then(|r| r.first()).cloned() {
                            self.set_pasted_format(*data_row, *col, format, all_formats, &mut format_patches, cx);
                        }
                    }
                }
                self.end_batch_and_broadcast(cx);

                // Build values grid for provenance
                if !target_cells.is_empty() {
                    values_grid.push(vec![single_value.clone()]);
                }

                // Record with provenance
                if !changes.is_empty() || !comment_patches.is_empty() || !format_patches.is_empty() {
                    let data_start_row = self.row_view.view_to_data(start_row);
                    let provenance = MutationOp::Paste {
                        sheet: self.sheet(cx).id,
                        dst_row: data_start_row,
                        dst_col: start_col,
                        values: values_grid,
                        mode: PasteMode::Both,
                    }.to_provenance(&self.sheet(cx).name);

                    let sheet_index = self.sheet_index(cx);
                    let mut actions = vec![UndoAction::Values { sheet_index, changes }];
                    if !comment_patches.is_empty() {
                        self.wb_mut(cx, |wb| wb.bump_revision_for_structure());
                        actions.push(UndoAction::Comments { sheet_index, patches: comment_patches, description: "Paste comments".into() });
                    }
                    if !format_patches.is_empty() {
                        actions.push(UndoAction::Format { sheet_index, patches: format_patches, kind: FormatActionKind::PasteFormats, description: "Paste formats".into() });
                    }
                    self.record_action_with_provenance(cx, UndoAction::Group { actions, description: "Paste".into() }, Some(provenance));
                    self.bump_cells_rev();
                    self.is_modified = true;
                }

                self.clipboard_visual_range = None;
                self.status_message = Some(format!("Pasted to {} cells", target_cells.len()));
                self.maybe_smoke_recalc(cx);
                cx.notify();
                return;
            }

            // Standard paste (multi-cell clipboard or single cell to single selection)
            // When filtered, paste to consecutive visible rows
            let data_start_row = self.row_view.view_to_data(start_row);

            // Compute paste rectangle for split-merge guard and merge recreation
            let paste_rows = parsed_grid.as_ref().map_or(lines.len(), |g| g.len());
            let paste_cols = parsed_grid.as_ref().map_or_else(
                || lines.iter().map(|l| l.split('\t').count()).max().unwrap_or(1),
                |g| g.iter().map(|r| r.len()).max().unwrap_or(1),
            );
            let paste_max_row = (data_start_row + paste_rows).saturating_sub(1);
            let paste_max_col = (start_col + paste_cols).saturating_sub(1);

            if self.block_table_paste(start_row, start_col, paste_rows, paste_cols, cx) { return; }

            // Refuse the whole paste if any target cell is pivot output.
            if self.block_if_pivot(start_row, start_col, start_row + paste_rows - 1, paste_max_col, "paste", cx) {
                return;
            }

            // Block if paste would split a merged region
            if let Some((mr, mc)) = self.paste_would_split_merge(data_start_row, start_col, paste_max_row, paste_max_col, cx) {
                self.status_message = Some(format!(
                    "Cannot paste: would split merged cells at {}{}. Unmerge first.",
                    Self::col_to_letter(mc), mr + 1,
                ));
                cx.notify();
                return;
            }

            if !self.sheet(cx).tables().is_empty() {
                let mut values: Vec<Vec<String>> = parsed_grid.clone().unwrap_or_else(|| lines.iter().map(|l| l.split('\t').map(str::to_owned).collect()).collect());
                if is_internal {
                    for (ri, row) in values.iter_mut().enumerate() {
                        for (ci, value) in row.iter_mut().enumerate() {
                            if value.starts_with('=') { *value = self.adjust_copied_formula(value, ri, ci, data_start_row + ri, start_col + ci); }
                        }
                    }
                }
                let objects = is_internal && (self.internal_clipboard.as_ref().is_some_and(|ic| !ic.merges.is_empty() || ic.comments.iter().flatten().any(Option::is_some))
                    || self.sheet(cx).comments().any(|((r,c),_)| r >= data_start_row && r <= paste_max_row && c >= start_col && c <= paste_max_col));
                if self.paste_table_growth(data_start_row, start_col, &values, objects, cx) { return; }
            }

            // For filtered paste: find the starting visible index
            let visible_start_idx = if is_filtered {
                self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
            } else {
                None
            };

            // Parse tab-separated values and build values grid for provenance
            let mut values_grid: Vec<Vec<String>> = Vec::new();
            let mut end_data_row = data_start_row;
            let mut end_col = start_col;

            self.wb_mut(cx, |wb| wb.begin_batch());

            // Choose iteration source: pre-parsed CSV grid or raw tab-split lines
            let row_count = parsed_grid.as_ref().map_or(lines.len(), |g| g.len());
            for row_offset in 0..row_count {
                // Determine target view row for this clipboard row
                let (_target_view_row, target_data_row) = if is_filtered {
                    if let Some(start_idx) = visible_start_idx {
                        // Get the nth visible row from the starting position
                        if let Some(view_row) = self.row_view.nth_visible(start_idx + row_offset) {
                            let data_row = self.row_view.view_to_data(view_row);
                            (view_row, data_row)
                        } else {
                            // No more visible rows - skip this line
                            continue;
                        }
                    } else {
                        // Start row not visible - skip
                        continue;
                    }
                } else {
                    // No filtering - direct mapping
                    let view_row = start_row + row_offset;
                    if view_row >= NUM_ROWS {
                        continue;
                    }
                    (view_row, view_row)
                };

                // Get columns for this row from either the parsed grid or tab-split
                let row_cells: Vec<&str> = if let Some(ref grid) = parsed_grid {
                    grid[row_offset].iter().map(|s| s.as_str()).collect()
                } else {
                    lines[row_offset].split('\t').collect()
                };

                let mut row_values: Vec<String> = Vec::new();
                for (col_offset, value) in row_cells.iter().enumerate() {
                    let col = start_col + col_offset;
                    if target_data_row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet(cx).get_raw(target_data_row, col);

                        // Adjust formula references using constant delta from source to destination
                        let new_value = if value.starts_with('=') && is_internal {
                            self.adjust_copied_formula(value, row_offset, col_offset, target_data_row, col)
                        } else {
                            value.to_string()
                        };

                        row_values.push(new_value.clone());

                        if old_value != new_value {
                            changes.push(CellChange {
                                row: target_data_row, col, old_value, new_value: new_value.clone(),
                            });
                        }
                        self.set_cell_value(target_data_row, col, &new_value, cx);
                        if is_internal {
                            let after = self.internal_clipboard.as_ref().and_then(|ic| ic.comments.get(row_offset)).and_then(|r| r.get(col_offset)).cloned().flatten();
                            let before = self.sheet(cx).comment(target_data_row, col).cloned();
                            if before != after {
                                self.active_sheet_mut(cx, |s| s.set_comment(target_data_row, col, after.clone()));
                                comment_patches.push(CommentPatch { remove_cell_on_undo: false, row: target_data_row, col, before, after });
                            }
                        }
                        if with_formats {
                            if let Some(format) = self.internal_clipboard.as_ref().and_then(|ic| ic.formats.get(row_offset)).and_then(|r| r.get(col_offset)).cloned() {
                                self.set_pasted_format(target_data_row, col, format, all_formats, &mut format_patches, cx);
                            }
                        }


                        // Track paste bounds (in data coordinates)
                        end_data_row = end_data_row.max(target_data_row);
                        end_col = end_col.max(col);
                    }
                }
                if !row_values.is_empty() {
                    values_grid.push(row_values);
                }
            }

            // Recreate clipboard merges at destination (only for internal paste, not filtered)
            let clipboard_merges = if is_internal {
                self.internal_clipboard.as_ref()
                    .map(|ic| ic.merges.clone())
                    .unwrap_or_default()
            } else {
                Vec::new()
            };

            let mut merge_action = None;
            if !clipboard_merges.is_empty() && !is_filtered {
                let merges_before = self.sheet(cx).merged_regions.clone();

                // Remove existing merges fully within the paste rectangle
                let origins_to_remove: Vec<(usize, usize)> = {
                    let sheet = self.sheet(cx);
                    sheet.merged_regions.iter()
                        .filter(|existing| {
                            existing.start.0 >= data_start_row
                            && existing.end.0 <= paste_max_row
                            && existing.start.1 >= start_col
                            && existing.end.1 <= paste_max_col
                        })
                        .map(|existing| existing.start)
                        .collect()
                };
                for origin in origins_to_remove {
                    self.active_sheet_mut(cx, |s| { let _ = s.remove_merge(origin); });
                }

                // Add clipboard merges at destination offsets
                let mut cleared_values: Vec<(usize, usize, String)> = Vec::new();
                for rel_merge in &clipboard_merges {
                    let dest = MergedRegion::new(
                        data_start_row + rel_merge.start.0,
                        start_col + rel_merge.start.1,
                        data_start_row + rel_merge.end.0,
                        start_col + rel_merge.end.1,
                    );
                    // Clear non-origin cells (same semantics as merge_cells)
                    for r in dest.start.0..=dest.end.0 {
                        for c in dest.start.1..=dest.end.1 {
                            if (r, c) == dest.start { continue; }
                            let raw = self.sheet(cx).get_raw(r, c);
                            if !raw.is_empty() {
                                cleared_values.push((r, c, raw));
                                self.set_cell_value(r, c, "", cx);
                            }
                        }
                    }
                    self.active_sheet_mut(cx, |s| { let _ = s.add_merge(dest); });
                }

                let merges_after = self.sheet(cx).merged_regions.clone();

                merge_action = Some(UndoAction::SetMerges {
                    sheet_index: self.sheet_index(cx),
                    before: merges_before,
                    after: merges_after,
                    cleared_values,
                    description: "Paste: recreate merges".to_string(),
                });
            }
            self.end_batch_and_broadcast(cx);

            // Record with provenance (only if changes or merge changes were made)
            if !changes.is_empty() || merge_action.is_some() || !comment_patches.is_empty() || !format_patches.is_empty() {
                let provenance = MutationOp::Paste {
                    sheet: self.sheet(cx).id,
                    dst_row: data_start_row,
                    dst_col: start_col,
                    values: values_grid,
                    mode: PasteMode::Both,
                }.to_provenance(&self.sheet(cx).name);

                let sheet_index = self.sheet_index(cx);
                let mut actions = vec![UndoAction::Values { sheet_index, changes }];
                if let Some(action) = merge_action { actions.push(action); }
                if !comment_patches.is_empty() {
                    self.wb_mut(cx, |wb| wb.bump_revision_for_structure());
                    actions.push(UndoAction::Comments { sheet_index, patches: comment_patches, description: "Paste comments".into() });
                }
                if !format_patches.is_empty() {
                    actions.push(UndoAction::Format { sheet_index, patches: format_patches, kind: FormatActionKind::PasteFormats, description: "Paste formats".into() });
                }
                self.record_action_with_provenance(cx, UndoAction::Group { actions, description: "Paste".into() }, Some(provenance));
                self.bump_cells_rev();
                self.is_modified = true;
            }

            // Validate pasted range and report failures (using data coordinates)
            let failures = self.wb(cx).validate_range(
                self.sheet_index(cx), data_start_row, start_col, end_data_row, end_col
            );
            let total_cells = (end_data_row - data_start_row + 1) * (end_col - start_col + 1);
            if failures.count > 0 {
                self.store_validation_failures(&failures);
                self.status_message = Some(format!(
                    "Pasted from clipboard (Validation: {} of {} cells failed) — Press F8 to jump",
                    failures.count, total_cells
                ));
            } else {
                self.status_message = Some("Pasted from clipboard".to_string());
            }

            // Clear copy border overlay — clipboard consumed
            self.clipboard_visual_range = None;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc(cx);

            cx.notify();
        }
    }

    /// Normalize clipboard text for comparison (handles line ending differences)
    pub(crate) fn normalize_clipboard_text(text: &str) -> String {
        // Normalize line endings and trim whitespace from both ends
        // Some clipboard managers add leading/trailing whitespace or transform line endings
        text.replace("\r\n", "\n").replace('\r', "\n").trim().to_string()
    }

    /// Determine if a paste operation should use internal clipboard data with formula adjustment.
    ///
    /// Returns true (internal paste) when:
    /// 1. Clipboard metadata matches internal clipboard ID (reliable cross-platform)
    /// 2. System clipboard text matches internal clipboard text (fallback when metadata unavailable
    ///    or when metadata doesn't match but text does — Wayland may garble metadata)
    /// 3. System clipboard is unavailable AND copy happened recently (< 2s, Wayland failure mode)
    ///
    /// Returns false (external paste) when:
    /// - No internal clipboard exists
    /// - System clipboard has different content (user copied from external source)
    ///
    /// This function is public for testing the Wayland clipboard-unavailable scenario.
    pub fn is_internal_paste(
        internal_clipboard: Option<&InternalClipboard>,
        system_text: Option<&str>,
        metadata: Option<&str>,
    ) -> bool {
        let Some(ic) = internal_clipboard else {
            return false;
        };

        let expected_id = format!("\"{}\"", ic.id);

        // Metadata match: definitive yes
        if let Some(m) = metadata {
            if m == expected_id {
                return true;
            }
            // Metadata exists but doesn't match — likely external.
            // Still check text as defensive fallback (metadata may be garbled on Wayland).
            if let Some(st) = system_text {
                return Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv);
            }
            return false;
        }

        // No metadata: fall back to text comparison
        if let Some(st) = system_text {
            return Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv);
        }

        // System clipboard completely unavailable (Wayland failure mode).
        // Only assume internal if the copy happened recently (< 2s).
        ic.created_at.elapsed() < std::time::Duration::from_secs(2)
    }

    /// Paste clipboard text into the edit buffer (when in editing mode)
    pub fn paste_into_edit(&mut self, cx: &mut Context<Self>) {
        let text = if let Some(item) = cx.read_from_clipboard() {
            item.text().map(|s| s.to_string())
        } else {
            self.internal_clipboard.as_ref().map(|ic| ic.raw_tsv.clone())
        };

        if let Some(text) = text {
            // Only take first line if multi-line, and trim whitespace
            let text = text.lines().next().unwrap_or("").trim();
            if !text.is_empty() {
                // Insert at cursor byte position
                let byte_pos = self.edit_cursor.min(self.edit_value.len());
                self.edit_value.insert_str(byte_pos, text);
                self.edit_cursor = byte_pos + text.len();  // Advance by byte length

                // Update autocomplete for formulas
                self.update_autocomplete(cx);

                self.edit_scroll_dirty = true;
                self.status_message = Some(format!("Pasted: {}", text));
                cx.notify();
            }
        }
    }

    /// Toggle the "Ctrl+V pastes values" default and persist it.
    pub fn toggle_paste_values_default(&mut self, cx: &mut Context<Self>) {
        let new_value = !crate::settings::user_settings(cx)
            .editing
            .paste_values_by_default
            .resolve(false);
        crate::settings::update_user_settings(cx, |s| {
            s.editing.paste_values_by_default = crate::settings::Setting::Value(new_value);
        });
        self.status_message = Some(if new_value {
            "Ctrl+V now pastes values only (full paste: Paste Special > All)".to_string()
        } else {
            "Ctrl+V now pastes everything (values only: Ctrl+Shift+V)".to_string()
        });
        cx.notify();
    }

    /// Paste Values: paste computed values only (no formulas).
    /// Uses typed values from internal clipboard, or parses external clipboard with leading-zero guard.
    /// When filtered, pastes to consecutive visible rows only.
    pub fn paste_values(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        if self.paste_table_headers(TablePasteKind::Values, cx) { return; }
        // Block during preview mode
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.paste_table_view(TablePasteKind::Values, cx);
            return;
        }
        if self.block_if_previewing(cx) { return; }

        // If editing, paste canonical text into edit buffer (top-left cell only)
        if self.mode.is_editing() {
            self.paste_values_into_edit(cx);
            return;
        }

        // Read clipboard item to get text
        let clipboard_item = cx.read_from_clipboard();
        let system_text = clipboard_item.as_ref().and_then(|item| item.text().map(|s| s.to_string()));

        // For Paste Values, prefer internal clipboard values if they exist and text matches.
        // This avoids depending on metadata (which doesn't round-trip on Windows).
        // The internal clipboard stores computed values, which is exactly what we want.
        let use_internal_values = self.internal_clipboard.as_ref().map_or(false, |ic| {
            // Use internal values if we have them AND either:
            // 1. System clipboard matches our raw_tsv (we copied it)
            // 2. System clipboard is empty/unavailable (use what we have)
            system_text.as_ref().map_or(true, |st| {
                Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv)
            })
        });

        let (start_row, start_col) = self.view_state.selected;
        let is_filtered = self.row_view.is_filtered();
        let data_start_row = self.row_view.view_to_data(start_row);

        // Block if paste would split a merged region
        {
            let (paste_rows, paste_cols) = if use_internal_values {
                self.internal_clipboard.as_ref()
                    .map(|ic| (ic.values.len(), ic.values.first().map_or(0, |r| r.len())))
                    .unwrap_or((0, 0))
            } else {
                let text = system_text.as_deref().unwrap_or("");
                let lines: Vec<&str> = text.lines().collect();
                (lines.len(), lines.iter().map(|l| l.split('\t').count()).max().unwrap_or(1))
            };
            if self.block_table_paste(start_row, start_col, paste_rows, paste_cols, cx) { return; }
            if paste_rows > 0 && paste_cols > 0 {
                let dest_max_row = (data_start_row + paste_rows).saturating_sub(1);
                let dest_max_col = (start_col + paste_cols).saturating_sub(1);
                if let Some((mr, mc)) = self.paste_would_split_merge(data_start_row, start_col, dest_max_row, dest_max_col, cx) {
                    self.status_message = Some(format!(
                        "Cannot paste: would split merged cells at {}{}. Unmerge first.",
                        Self::col_to_letter(mc), mr + 1,
                    ));
                    cx.notify();
                    return;
                }
            }
        }

        if !self.sheet(cx).tables().is_empty() {
            let values: Vec<Vec<String>> = if use_internal_values {
                self.internal_clipboard.as_ref().map(|ic| ic.values.iter().map(|row| row.iter().map(Self::value_to_canonical_string).collect()).collect()).unwrap_or_default()
            } else {
                system_text.as_deref().unwrap_or("").lines().map(|line| line.split('\t').map(|value| Self::value_to_canonical_string(&Self::parse_external_value(value))).collect()).collect()
            };
            if self.paste_table_growth(data_start_row, start_col, &values, false, cx) { return; }
        }

        let mut changes = Vec::new();
        let mut values_grid: Vec<Vec<String>> = Vec::new();
        let mut end_data_row = data_start_row;
        let mut end_col = start_col;

        // For filtered paste: find the starting visible index
        let visible_start_idx = if is_filtered {
            self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
        } else {
            None
        };

        /// Helper to get the nth target row (returns data_row)
        fn get_target_data_row(
            row_view: &visigrid_engine::filter::RowView,
            is_filtered: bool,
            visible_start_idx: Option<usize>,
            start_row: usize,
            row_offset: usize,
        ) -> Option<usize> {
            if is_filtered {
                if let Some(start_idx) = visible_start_idx {
                    if let Some(view_row) = row_view.nth_visible(start_idx + row_offset) {
                        return Some(row_view.view_to_data(view_row));
                    }
                }
                None
            } else {
                let view_row = start_row + row_offset;
                if view_row < NUM_ROWS {
                    Some(view_row)
                } else {
                    None
                }
            }
        }

        self.wb_mut(cx, |wb| wb.begin_batch());
        if use_internal_values {
            // Use typed values from internal clipboard (clone to avoid borrow issues)
            let values = self.internal_clipboard.as_ref().map(|ic| ic.values.clone());
            if let Some(values) = values {
                for (row_offset, row_values) in values.iter().enumerate() {
                    let Some(target_data_row) = get_target_data_row(
                        &self.row_view, is_filtered, visible_start_idx, start_row, row_offset
                    ) else {
                        continue;
                    };

                    let mut grid_row: Vec<String> = Vec::new();
                    for (col_offset, value) in row_values.iter().enumerate() {
                        let col = start_col + col_offset;
                        if target_data_row < NUM_ROWS && col < NUM_COLS {
                            let old_value = self.sheet(cx).get_raw(target_data_row, col);
                            let new_value = Self::value_to_canonical_string(value);

                            grid_row.push(new_value.clone());

                            if old_value != new_value {
                                changes.push(CellChange {
                                    row: target_data_row, col, old_value, new_value: new_value.clone(),
                                });
                            }
                            self.set_cell_value(target_data_row, col, &new_value, cx);

                            end_data_row = end_data_row.max(target_data_row);
                            end_col = end_col.max(col);
                        }
                    }
                    if !grid_row.is_empty() {
                        values_grid.push(grid_row);
                    }
                }
            }
        } else if let Some(text) = system_text {
            // Parse external clipboard with leading-zero guard
            for (row_offset, line) in text.lines().enumerate() {
                let Some(target_data_row) = get_target_data_row(
                    &self.row_view, is_filtered, visible_start_idx, start_row, row_offset
                ) else {
                    continue;
                };

                let mut grid_row: Vec<String> = Vec::new();
                for (col_offset, cell_text) in line.split('\t').enumerate() {
                    let col = start_col + col_offset;
                    if target_data_row < NUM_ROWS && col < NUM_COLS {
                        let old_value = self.sheet(cx).get_raw(target_data_row, col);
                        let parsed_value = Self::parse_external_value(cell_text);
                        let new_value = Self::value_to_canonical_string(&parsed_value);

                        grid_row.push(new_value.clone());

                        if old_value != new_value {
                            changes.push(CellChange {
                                row: target_data_row, col, old_value, new_value: new_value.clone(),
                            });
                        }
                        self.set_cell_value(target_data_row, col, &new_value, cx);

                        end_data_row = end_data_row.max(target_data_row);
                        end_col = end_col.max(col);
                    }
                }
                if !grid_row.is_empty() {
                    values_grid.push(grid_row);
                }
            }
        }
        self.end_batch_and_broadcast(cx);

        if !changes.is_empty() {
            let provenance = MutationOp::Paste {
                sheet: self.sheet(cx).id,
                dst_row: data_start_row,
                dst_col: start_col,
                values: values_grid,
                mode: PasteMode::Values,
            }.to_provenance(&self.sheet(cx).name);

            self.record_batch_with_provenance(cx, self.sheet_index(cx), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;

            // Smoke mode: trigger full ordered recompute for dogfooding
            self.maybe_smoke_recalc(cx);
        }

        // Validate pasted range and report failures (using data coordinates)
        let failures = self.wb(cx).validate_range(
            self.sheet_index(cx), data_start_row, start_col, end_data_row, end_col
        );
        let total_cells = (end_data_row - data_start_row + 1) * (end_col - start_col + 1);
        if failures.count > 0 {
            self.store_validation_failures(&failures);
            self.status_message = Some(format!(
                "Pasted values (Validation: {} of {} cells failed) — Press F8 to jump",
                failures.count, total_cells
            ));
        } else {
            self.status_message = Some("Pasted values".to_string());
        }
        self.clipboard_visual_range = None;
        cx.notify();
    }

    /// Convert a typed Value to its canonical string representation for cell storage.
    /// Guarantees: no scientific notation, deterministic output, -0.0 normalized to 0.
    pub(crate) fn value_to_canonical_string(value: &Value) -> String {
        match value {
            Value::Empty => String::new(),
            Value::Number(n) => {
                // Handle non-finite values explicitly
                if !n.is_finite() {
                    if n.is_nan() { return "NaN".to_string(); }
                    return if *n > 0.0 { "INF".to_string() } else { "-INF".to_string() };
                }

                // Normalize -0.0 to 0.0
                let n0 = if *n == 0.0 { 0.0 } else { *n };

                // Integer fast path: no decimal point needed
                if n0.fract() == 0.0 && n0.abs() < 9e15 {
                    format!("{:.0}", n0)
                } else {
                    // Fixed precision (15 decimals), trim trailing zeros, no scientific notation
                    let mut s = format!("{:.15}", n0);
                    while s.contains('.') && s.ends_with('0') { s.pop(); }
                    if s.ends_with('.') { s.pop(); }
                    s
                }
            }
            Value::Text(s) => s.clone(),
            Value::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
            Value::Error(e) => e.clone(),
        }
    }

    /// Parse external clipboard text into a typed Value with leading-zero preservation.
    pub(crate) fn parse_external_value(text: &str) -> Value {
        let trimmed = text.trim();

        if trimmed.is_empty() {
            return Value::Empty;
        }

        // Check for formula prefix - treat as literal text (strip the =)
        if trimmed.starts_with('=') {
            return Value::Text(trimmed.to_string());
        }

        // Check for leading zeros that should be preserved as text
        // e.g., "007", "00123" - but not "0" or "0.5"
        if trimmed.starts_with('0') && trimmed.len() > 1 {
            let second_char = trimmed.chars().nth(1).unwrap();
            if second_char.is_ascii_digit() {
                // Starts with 0 followed by digit -> preserve as text
                return Value::Text(trimmed.to_string());
            }
        }

        // Check for boolean
        let upper = trimmed.to_uppercase();
        if upper == "TRUE" {
            return Value::Boolean(true);
        }
        if upper == "FALSE" {
            return Value::Boolean(false);
        }

        // Try to parse as formatted number (commas, currency, parens)
        if let Some(n) = visigrid_engine::cell::try_parse_number(trimmed) {
            return Value::Number(n);
        }

        // Try to parse as plain number
        if let Ok(n) = trimmed.parse::<f64>() {
            return Value::Number(n);
        }

        // Default to text
        Value::Text(trimmed.to_string())
    }

    /// Paste values into edit buffer: use canonical text of top-left value only.
    fn paste_values_into_edit(&mut self, cx: &mut Context<Self>) {
        // Read clipboard item to get text
        let clipboard_item = cx.read_from_clipboard();
        let system_text = clipboard_item.as_ref().and_then(|item| item.text().map(|s| s.to_string()));

        // For Paste Values, prefer internal clipboard values if they exist and text matches.
        // This avoids depending on metadata (which doesn't round-trip on Windows).
        let use_internal_values = self.internal_clipboard.as_ref().map_or(false, |ic| {
            system_text.as_ref().map_or(true, |st| {
                Self::normalize_clipboard_text(st) == Self::normalize_clipboard_text(&ic.raw_tsv)
            })
        });

        let text = if use_internal_values {
            // Get top-left value from internal clipboard
            self.internal_clipboard.as_ref().and_then(|ic| {
                ic.values.first().and_then(|row| row.first()).map(|v| Self::value_to_canonical_string(v))
            })
        } else {
            // Parse top-left cell from external clipboard
            system_text.map(|text| {
                let first_cell = text.lines().next().unwrap_or("")
                    .split('\t').next().unwrap_or("");
                let value = Self::parse_external_value(first_cell);
                Self::value_to_canonical_string(&value)
            })
        };

        if let Some(text) = text {
            if !text.is_empty() {
                // Insert at cursor byte position
                let byte_pos = self.edit_cursor.min(self.edit_value.len());
                self.edit_value.insert_str(byte_pos, &text);
                self.edit_cursor = byte_pos + text.len();  // Advance by byte length

                self.update_autocomplete(cx);
                self.edit_scroll_dirty = true;
                self.status_message = Some(format!("Pasted value: {}", text));
                cx.notify();
            }
        }
    }

    /// Paste Formulas: paste raw formulas with reference adjustment.
    /// - Internal clipboard: uses raw_tsv (formulas) with reference adjustment
    /// - External clipboard: falls back to normal paste() (no way to distinguish formula vs text)
    pub fn paste_formulas(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.paste_table_view(TablePasteKind::Formulas, cx);
            return;
        }
        if self.block_if_previewing(cx) { return; }

        // If editing, paste into edit buffer
        if self.mode.is_editing() {
            self.paste_into_edit(cx);
            return;
        }

        // Check if we have an internal clipboard with matching ID
        let clipboard_item = cx.read_from_clipboard();
        let metadata = clipboard_item.as_ref().and_then(|item| item.metadata().cloned());

        let is_internal = self.internal_clipboard.as_ref().map_or(false, |ic| {
            let expected_id = format!("\"{}\"", ic.id);
            metadata.as_ref().map_or(false, |m| m == &expected_id)
        });

        if !is_internal {
            // External clipboard - fall back to normal paste
            // (No way to reliably distinguish "formula" vs "text starting with =" from external)
            self.paste(cx);
            return;
        }

        // Internal paste - use raw_tsv with reference adjustment
        let (start_row, start_col) = self.view_state.selected;
        let is_filtered = self.row_view.is_filtered();
        let data_start_row = self.row_view.view_to_data(start_row);

        let raw_cells = self.internal_clipboard.as_ref().map(|ic| ic.raw_cells.clone()).unwrap_or_default();
        // Block if paste would split a merged region
        {
            let paste_rows = raw_cells.len();
            let paste_cols = raw_cells.iter().map(Vec::len).max().unwrap_or(1);
            if self.block_table_paste(start_row, start_col, paste_rows, paste_cols, cx) { return; }
            if paste_rows > 0 && paste_cols > 0 {
                let dest_max_row = (data_start_row + paste_rows).saturating_sub(1);
                let dest_max_col = (start_col + paste_cols).saturating_sub(1);
                if let Some((mr, mc)) = self.paste_would_split_merge(data_start_row, start_col, dest_max_row, dest_max_col, cx) {
                    self.status_message = Some(format!(
                        "Cannot paste: would split merged cells at {}{}. Unmerge first.",
                        Self::col_to_letter(mc), mr + 1,
                    ));
                    cx.notify();
                    return;
                }
            }
        }

        let mut changes = Vec::new();
        let mut values_grid: Vec<Vec<String>> = Vec::new();
        let mut end_data_row = data_start_row;
        let mut end_col = start_col;

        if !self.sheet(cx).tables().is_empty() {
            let values: Vec<Vec<String>> = raw_cells.iter().enumerate().map(|(ri,row)| row.iter().enumerate().map(|(ci,value)| {
                if value.starts_with('=') { self.adjust_copied_formula(value, ri, ci, data_start_row + ri, start_col + ci) } else { value.to_owned() }
            }).collect()).collect();
            if self.paste_table_growth(data_start_row, start_col, &values, false, cx) { return; }
        }

        // For filtered paste: find the starting visible index
        let visible_start_idx = if is_filtered {
            self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
        } else {
            None
        };

        self.wb_mut(cx, |wb| wb.begin_batch());
        for (row_offset, row) in raw_cells.iter().enumerate() {
            // Determine target view row for this clipboard row
            let target_data_row = if is_filtered {
                if let Some(start_idx) = visible_start_idx {
                    if let Some(view_row) = self.row_view.nth_visible(start_idx + row_offset) {
                        self.row_view.view_to_data(view_row)
                    } else {
                        continue;
                    }
                } else {
                    continue;
                }
            } else {
                let view_row = start_row + row_offset;
                if view_row >= NUM_ROWS { continue; }
                view_row
            };

            let mut row_values: Vec<String> = Vec::new();
            for (col_offset, value) in row.iter().enumerate() {
                let col = start_col + col_offset;
                if target_data_row < NUM_ROWS && col < NUM_COLS {
                    let old_value = self.sheet(cx).get_raw(target_data_row, col);

                    // Adjust formula references using constant delta from source to destination
                    let new_value = if value.starts_with('=') {
                        self.adjust_copied_formula(value, row_offset, col_offset, target_data_row, col)
                    } else {
                        value.to_string()
                    };

                    row_values.push(new_value.clone());

                    if old_value != new_value {
                        changes.push(CellChange {
                            row: target_data_row, col, old_value, new_value: new_value.clone(),
                        });
                    }
                    self.set_cell_value(target_data_row, col, &new_value, cx);

                    end_data_row = end_data_row.max(target_data_row);
                    end_col = end_col.max(col);
                }
            }
            if !row_values.is_empty() {
                values_grid.push(row_values);
            }
        }
        self.end_batch_and_broadcast(cx);

        // Record with provenance (PasteMode::Formulas)
        if !changes.is_empty() {
            let provenance = MutationOp::Paste {
                sheet: self.sheet(cx).id,
                dst_row: data_start_row,
                dst_col: start_col,
                values: values_grid,
                mode: PasteMode::Formulas,
            }.to_provenance(&self.sheet(cx).name);

            self.record_batch_with_provenance(cx, self.sheet_index(cx), changes, Some(provenance));
            self.bump_cells_rev();
            self.is_modified = true;
        }

        self.clipboard_visual_range = None;
        self.status_message = Some("Pasted formulas".to_string());
        self.maybe_smoke_recalc(cx);
        cx.notify();
    }

    /// Paste Formats: paste cell formatting only (no values).
    /// - Internal clipboard only: applies formats from copied range
    /// - External clipboard: no-op with status message (no format data available)
    pub fn paste_formats(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.paste_table_view(TablePasteKind::Formats, cx);
            return;
        }
        if self.block_if_previewing(cx) { return; }

        // Paste Formats doesn't make sense in edit mode
        if self.mode.is_editing() {
            self.status_message = Some("Exit edit mode to paste formats".to_string());
            cx.notify();
            return;
        }

        // Check if we have an internal clipboard with matching ID
        let clipboard_item = cx.read_from_clipboard();
        let metadata = clipboard_item.as_ref().and_then(|item| item.metadata().cloned());

        let is_internal = self.internal_clipboard.as_ref().map_or(false, |ic| {
            let expected_id = format!("\"{}\"", ic.id);
            metadata.as_ref().map_or(false, |m| m == &expected_id)
        });

        if !is_internal {
            // External clipboard - no format data available
            self.status_message = Some("Paste Formats requires VisiGrid clipboard".to_string());
            cx.notify();
            return;
        }

        // Get formats from internal clipboard
        let formats = match &self.internal_clipboard {
            Some(ic) if !ic.formats.is_empty() => ic.formats.clone(),
            _ => {
                self.status_message = Some("No formats in clipboard".to_string());
                cx.notify();
                return;
            }
        };

        let (start_row, start_col) = self.view_state.selected;
        let is_filtered = self.row_view.is_filtered();
        let data_start_row = self.row_view.view_to_data(start_row);

        // For filtered paste: find the starting visible index
        let visible_start_idx = if is_filtered {
            self.row_view.visible_rows().iter().position(|&vr| vr == start_row)
        } else {
            None
        };

        let mut format_patches = Vec::new();

        for (row_offset, row_formats) in formats.iter().enumerate() {
            // Determine target data row
            let target_data_row = if is_filtered {
                if let Some(start_idx) = visible_start_idx {
                    if let Some(view_row) = self.row_view.nth_visible(start_idx + row_offset) {
                        self.row_view.view_to_data(view_row)
                    } else {
                        continue;
                    }
                } else {
                    continue;
                }
            } else {
                let view_row = start_row + row_offset;
                if view_row >= NUM_ROWS { continue; }
                view_row
            };

            for (col_offset, format) in row_formats.iter().enumerate() {
                let col = start_col + col_offset;
                if target_data_row < NUM_ROWS && col < NUM_COLS {
                    // Get old format for history
                    let old_format = self.sheet(cx).get_format(target_data_row, col).clone();

                    // Apply format (full replace)
                    self.active_sheet_mut(cx, |s| {
                        s.set_format(target_data_row, col, format.clone());
                    });

                    // Track change for history
                    if old_format != *format {
                        format_patches.push(CellFormatPatch { remove_cell_on_undo: false,
                            row: target_data_row,
                            col,
                            before: old_format,
                            after: format.clone(),
                        });
                    }
                }
            }
        }

        // Record format changes in history with provenance
        if !format_patches.is_empty() {
            let provenance = MutationOp::Paste {
                sheet: self.sheet(cx).id,
                dst_row: data_start_row,
                dst_col: start_col,
                values: vec![], // No value changes
                mode: PasteMode::Formats,
            }.to_provenance(&self.sheet(cx).name);

            self.record_format_with_provenance(cx,
                self.sheet_index(cx),
                format_patches,
                FormatActionKind::PasteFormats,
                "Paste Formats".to_string(),
                Some(provenance),
            );
            self.is_modified = true;
        }

        self.clipboard_visual_range = None;
        let rows = formats.len();
        let cols = formats.first().map(|r| r.len()).unwrap_or(0);
        self.status_message = Some(format!("Pasted formats to {}x{} range", rows, cols));
        cx.notify();
    }

    pub fn delete_selection(&mut self, cx: &mut Context<Self>) {
        // Block during preview mode
        if crate::table_filter_ui::has_table_criteria(self.wb(cx)) {
            self.delete_table_selection(cx);
            return;
        }
        if self.block_if_previewing(cx) { return; }
        if self.block_if_selection_in_pivot("clear", cx) { return; }
        if self.block_selection_table_headers("clear", cx) { return; }

        let mut changes = Vec::new();
        let mut skipped_spill_receivers = false;
        let is_filtered = self.row_view.is_filtered();

        // Delete from all selection ranges (including discontiguous Ctrl+Click selections)
        // When filtered, only delete from visible rows
        self.wb_mut(cx, |wb| wb.begin_batch());
        let rows_are_identity = !is_filtered && !self.row_view.is_sorted();
        for ((min_row, min_col), (max_row, max_col)) in self.all_selection_ranges() {
            if rows_are_identity {
                // Sparse fast path: enumerate only populated cells intersecting
                // the range. Ctrl+A selects 16.7M coordinates; walking them all
                // beachballs for seconds while the sheet holds a few dozen
                // values. Excel's clear is O(actual data) — so is this.
                let mut targets: Vec<(usize, usize)> = self
                    .sheet(cx)
                    .cells_iter()
                    .filter(|((r, c), _)| {
                        *r >= min_row && *r <= max_row && *c >= min_col && *c <= max_col
                    })
                    .map(|((r, c), _)| (r, c))
                    .collect();
                targets.sort_unstable(); // deterministic change/undo order

                for (data_row, col) in targets {
                    if self.sheet(cx).is_spill_receiver(data_row, col) {
                        skipped_spill_receivers = true;
                        continue;
                    }
                    let old_value = self.sheet(cx).get_raw(data_row, col);
                    if !old_value.is_empty() {
                        changes.push(CellChange {
                            row: data_row, col, old_value, new_value: String::new(),
                        });
                    }
                    self.clear_cell_value(data_row, col, cx);
                }
                continue;
            }

            // Sorted/filtered views: selection rows are view-space, so walk the
            // rectangle and map each row (bounded by the visible row count in
            // practice, since filtered/sorted views operate on real data extents)
            for view_row in min_row..=max_row {
                // Skip hidden rows when filtered
                if is_filtered && !self.row_view.is_view_row_visible(view_row) {
                    continue;
                }

                // Convert view row to data row for sheet access
                let data_row = self.row_view.view_to_data(view_row);

                for col in min_col..=max_col {
                    // Skip spill receivers - only the parent formula can be deleted
                    if self.sheet(cx).is_spill_receiver(data_row, col) {
                        skipped_spill_receivers = true;
                        continue;
                    }

                    let old_value = self.sheet(cx).get_raw(data_row, col);
                    if !old_value.is_empty() {
                        changes.push(CellChange {
                            row: data_row, col, old_value, new_value: String::new(),
                        });
                    }
                    self.clear_cell_value(data_row, col, cx);
                }
            }
        }
        self.end_batch_and_broadcast(cx);

        let had_changes = !changes.is_empty();
        if had_changes {
            // Only attach provenance for single contiguous selection
            let provenance = if self.view_state.additional_selections.is_empty() {
                let ((min_row, min_col), (max_row, max_col)) = self.selection_range();
                // Use data rows for provenance
                let data_min_row = self.row_view.view_to_data(min_row);
                let data_max_row = self.row_view.view_to_data(max_row);
                Some(MutationOp::Clear {
                    sheet: self.sheet(cx).id,
                    start_row: data_min_row,
                    start_col: min_col,
                    end_row: data_max_row,
                    end_col: max_col,
                    mode: ClearMode::All,
                }.to_provenance(&self.sheet(cx).name))
            } else {
                None  // Discontiguous selection - no provenance
            };
            self.record_batch_with_provenance(cx, self.sheet_index(cx), changes, provenance);
            self.bump_cells_rev();  // Invalidate cell search cache
            self.is_modified = true;
        }

        if had_changes {
            self.clipboard_visual_range = None;
        }

        if skipped_spill_receivers && !had_changes {
            self.status_message = Some("Cannot delete spill range. Delete the parent formula instead.".to_string());
        }

        cx.notify();
    }

    /// Check if pasting into (min_row..=max_row, min_col..=max_col) would split any merge.
    /// Returns Some(merge_origin) for the first offending merge, or None if safe.
    fn paste_would_split_merge(
        &self, min_row: usize, min_col: usize,
        max_row: usize, max_col: usize, cx: &App,
    ) -> Option<(usize, usize)> {
        let sheet = self.sheet(cx);
        if sheet.merged_regions.is_empty() { return None; }

        for merge in &sheet.merged_regions {
            let intersects = merge.end.0 >= min_row && merge.start.0 <= max_row
                          && merge.end.1 >= min_col && merge.start.1 <= max_col;
            if !intersects { continue; }

            let contained = merge.start.0 >= min_row && merge.end.0 <= max_row
                          && merge.start.1 >= min_col && merge.end.1 <= max_col;
            if !contained {
                return Some(merge.start);
            }
        }
        None
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum TablePasteKind {
    Contents,
    All,
    Values,
    Formulas,
    Formats,
}

impl Spreadsheet {
    fn adjust_copied_formula(
        &self,
        formula: &str,
        source_offset: usize,
        source_col: usize,
        row: usize,
        col: usize,
    ) -> String {
        let Some(ic) = &self.internal_clipboard else {
            return formula.to_owned();
        };
        let source_row = ic
            .source_rows
            .get(source_offset)
            .copied()
            .unwrap_or(ic.source.0 + source_offset);
        self.adjust_formula_refs(
            formula,
            row as i32 - source_row as i32,
            col as i32 - (ic.source.1 + source_col) as i32,
        )
    }

    fn paste_table_view(&mut self, kind: TablePasteKind, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) {
            return;
        }
        if self.mode.is_editing() {
            if kind == TablePasteKind::Values {
                self.paste_values_into_edit(cx);
            } else if kind != TablePasteKind::Formats {
                self.paste_into_edit(cx);
            }
            return;
        }
        self.sync_table_view(cx);
        let result = self.plan_table_paste(kind, cx);
        match result {
            Ok((writes, Some(plan))) => self.paste_and_append_table(plan, writes, cx),
            Ok((writes, None)) => {
                self.apply_table_cell_writes(writes, "Paste cells", cx);
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
            }
        }
    }

    fn plan_table_paste(
        &self,
        kind: TablePasteKind,
        cx: &App,
    ) -> Result<(Vec<crate::table_edit::TableCellWrite>, Option<crate::table_bulk_append::BulkAppendPlan>), String> {
        let item = cx.read_from_clipboard();
        let text = item.as_ref().and_then(|i| i.text());
        let metadata = item.as_ref().and_then(|i| i.metadata());
        let internal = Self::is_internal_paste(
            self.internal_clipboard.as_ref(),
            text.as_deref(),
            metadata.map(|s| s.as_str()),
        );
        let ic = if internal {
            self.internal_clipboard.as_ref()
        } else {
            None
        };
        if kind == TablePasteKind::Formats && ic.is_none() {
            return Err("Copy cells in VisiGrid before pasting formats.".into());
        }
        if ic.is_some_and(|ic| !ic.merges.is_empty()) {
            return Err(
                "Cannot paste merged cells through a Table view. Unmerge the source first.".into(),
            );
        }
        let text = ic
            .map(|ic| ic.raw_tsv.as_str())
            .or(text.as_deref())
            .ok_or("The clipboard is empty.")?;
        let grid: Vec<Vec<String>> = if let Some(ic) = ic {
            ic.raw_cells.clone()
        } else if !text.contains('\t')
            && !text
                .lines()
                .all(|line| visigrid_engine::cell::try_parse_number(line.trim()).is_some())
        {
            csv_io::parse_delimited_text(text)
        } else {
            full_paste_lines(text, internal)
                .iter()
                .map(|line| line.split('\t').map(str::to_owned).collect())
                .collect()
        };
        if grid.is_empty() {
            return Err("The clipboard is empty.".into());
        }
        let width = grid.iter().map(Vec::len).max().unwrap_or(0);
        if grid.len().saturating_mul(width) > 100_000 {
            return Err("Paste at most 100,000 cells at a time through a Table view.".into());
        }
        let (start, col) = self.view_state.selected;
        let broadcast = grid.len() == 1 && width == 1 && self.is_multi_selection();
        let mut append = None;
        let targets: Vec<(usize, usize, usize, usize)> = if broadcast {
            self.table_selection_targets(cx)?
                .into_iter()
                .map(|(r, c)| (r, c, 0, 0))
                .collect()
        } else {
            if !self.view_state.additional_selections.is_empty() {
                return Err("Select one destination for a multi-cell paste.".into());
            }
            if kind != TablePasteKind::Formats {
                append = crate::table_bulk_append::plan_bulk_append(
                    self.sheet(cx), &self.row_view, (start, col), grid.len(), width,
                )?;
            }
            if let Some(plan) = &append {
                plan.targets.clone()
            } else {
                crate::table_edit::view_safe_paste_targets(
                    self.sheet(cx), &self.row_view, (start, col), grid.len(), width,
                )?
            }
        };
        Ok((table_paste_writes_for_sheet(self.sheet(cx), &grid, ic, kind, targets), append))
    }
}

/// Match ordinary Ctrl+V: bring source formats only into unformatted cells.
/// Decide against canonical destinations before the atomic Table transaction.
fn table_paste_writes_for_sheet(
    sheet: &visigrid_engine::sheet::Sheet,
    grid: &[Vec<String>],
    ic: Option<&InternalClipboard>,
    kind: TablePasteKind,
    targets: Vec<(usize, usize, usize, usize)>,
) -> Vec<crate::table_edit::TableCellWrite> {
    let source_kind = if kind == TablePasteKind::Contents { TablePasteKind::All } else { kind };
    let mut writes = table_paste_writes(grid, ic, source_kind, targets);
    if kind == TablePasteKind::Contents {
        for write in &mut writes {
            if sheet.get_format(write.row, write.col) != CellFormat::default() {
                write.format = None;
            }
        }
    }
    writes
}

pub(crate) fn table_paste_writes(
    grid: &[Vec<String>],
    ic: Option<&InternalClipboard>,
    kind: TablePasteKind,
    targets: Vec<(usize, usize, usize, usize)>,
) -> Vec<crate::table_edit::TableCellWrite> {
    use crate::table_edit::TableCellWrite;
    let mut writes = Vec::with_capacity(targets.len());
    for (row, col, ri, ci) in targets {
        let raw = grid[ri].get(ci).map(String::as_str).unwrap_or("");
        let mut write = TableCellWrite::value(row, col, raw.to_owned());
        if kind == TablePasteKind::Formats {
            write.value = None;
        } else if kind == TablePasteKind::Values {
            let value = ic
                .and_then(|ic| ic.values.get(ri)?.get(ci))
                .cloned()
                .unwrap_or_else(|| Spreadsheet::parse_external_value(raw));
            write.literal_text = matches!(value, Value::Text(_));
            write.value = Some(Spreadsheet::value_to_canonical_string(&value));
        } else if let Some(ic) = ic {
            // Raw text that looks like a formula or number must remain text.
            write.literal_text = matches!(
                ic.values.get(ri).and_then(|r| r.get(ci)),
                Some(Value::Text(_))
            ) && !ic
                .source_formulas
                .get(ri)
                .and_then(|r| r.get(ci))
                .copied()
                .unwrap_or(false);
            if raw.starts_with('=') && !write.literal_text {
                let source_row = ic.source_rows.get(ri).copied().unwrap_or(ic.source.0 + ri);
                write.value = Some(visigrid_engine::formula::parser::adjust_formula_refs(
                    raw,
                    row as i32 - source_row as i32,
                    col as i32 - (ic.source.1 + ci) as i32,
                ));
            }
        }
        if let Some(ic) = ic {
            if matches!(kind, TablePasteKind::All | TablePasteKind::Formats) {
                write.format = ic.formats.get(ri).and_then(|r| r.get(ci)).cloned();
            }
            if matches!(kind, TablePasteKind::All | TablePasteKind::Contents) {
                write.comment = Some(
                    ic.comments
                        .get(ri)
                        .and_then(|r| r.get(ci))
                        .cloned()
                        .flatten(),
                );
            }
        }
        writes.push(write);
    }
    writes
}

#[cfg(test)]
mod table_paste_tests {
    use super::{table_paste_writes, InternalClipboard, TablePasteKind};
    use visigrid_engine::{cell::CellFormat, formula::eval::Value};

    fn clipboard() -> InternalClipboard {
        InternalClipboard {
            raw_tsv: "=A9\n=A4".into(),
            raw_cells: vec![vec!["=A9".into()], vec!["=A4".into()]],
            values: vec![vec![Value::Number(9.0)], vec![Value::Number(4.0)]],
            formats: vec![vec![CellFormat::default()]; 2],
            comments: vec![vec![None]; 2],
            source: (8, 1),
            source_rows: vec![8, 3],
            source_formulas: vec![vec![true]; 2],
            id: 1,
            merges: vec![],
            created_at: std::time::Instant::now(),
        }
    }

    #[test]
    fn formula_paste_uses_each_canonical_source_and_destination() {
        let ic = clipboard();
        let grid = vec![vec!["=(A9+$A$1)*2".into()], vec!["=$A4+A$1".into()]];
        let writes = table_paste_writes(
            &grid,
            Some(&ic),
            TablePasteKind::Contents,
            vec![(5, 2, 0, 0), (9, 2, 1, 0)],
        );
        assert_eq!(writes[0].value.as_deref(), Some("=(B6+$A$1)*2"));
        assert_eq!(writes[1].value.as_deref(), Some("=$A10+B$1"));
        // Broadcast has one source, regardless of destination gaps.
        let writes = table_paste_writes(
            &grid,
            Some(&ic),
            TablePasteKind::Formulas,
            vec![(5, 2, 0, 0), (9, 2, 0, 0)],
        );
        assert_eq!(writes[1].value.as_deref(), Some("=(B10+$A$1)*2"));
    }

    #[test]
    fn paste_values_preserves_formula_looking_text_and_leading_zeroes() {
        let grid = vec![vec!["=1+1".into(), "00123".into()]];
        let writes = table_paste_writes(
            &grid,
            None,
            TablePasteKind::Values,
            vec![(3, 1, 0, 0), (3, 2, 0, 1)],
        );
        assert!(writes.iter().all(|w| w.literal_text));
        assert_eq!(writes[0].value.as_deref(), Some("=1+1"));
        assert_eq!(writes[1].value.as_deref(), Some("00123"));
    }

    #[test]
    fn filtered_contents_paste_preserves_destination_formats_and_undo() {
        use crate::table_edit::{prepare_table_writes, tests::fixture, view_safe_paste_targets};
        let mut before = fixture(true);
        let italic = CellFormat { italic: true, ..Default::default() };
        before.sheet_mut(0).unwrap().set_format(3, 2, italic.clone());
        let mut ic = clipboard();
        ic.raw_cells = vec![vec!["25".into()], vec!["35".into()]];
        ic.values = vec![vec![Value::Number(25.0)], vec![Value::Number(35.0)]];
        ic.source_formulas = vec![vec![false]; 2];
        for formats in &mut ic.formats { formats[0].bold = true; }
        let sheet = before.active_sheet();
        let view = sheet.build_saved_table_view(30).unwrap().unwrap();
        // Slot 3 contains the filtered-out East record; slot 4 is the first visible record.
        let targets = view_safe_paste_targets(sheet, view.rows(), (4, 2), 2, 1).unwrap();
        assert_eq!(targets.iter().map(|t| t.0).collect::<Vec<_>>(), vec![5, 3]);
        let writes = super::table_paste_writes_for_sheet(
            sheet, &ic.raw_cells, Some(&ic), TablePasteKind::Contents, targets.clone());
        let mut after = prepare_table_writes(&before, 0, &writes).unwrap();
        assert!(after.active_sheet().get_format(5, 2).bold);
        assert_eq!(after.active_sheet().get_format(3, 2), italic);
        assert_eq!(after.active_sheet().get_raw(4, 2), "10");
        let commit = before.capture_guarded_batch(&after).unwrap();
        commit.replay(&mut after, true).unwrap();
        assert_eq!(after.active_sheet().get_format(5, 2), CellFormat::default());
        assert_eq!(after.active_sheet().get_format(3, 2), italic);
        commit.replay(&mut after, false).unwrap();
        assert!(after.active_sheet().get_format(5, 2).bold);
        let all = super::table_paste_writes_for_sheet(
            sheet, &ic.raw_cells, Some(&ic), TablePasteKind::All, targets);
        let all = prepare_table_writes(&before, 0, &all).unwrap();
        assert!(all.active_sheet().get_format(3, 2).bold);
        assert!(!all.active_sheet().get_format(3, 2).italic);
    }

    #[test]
    fn paste_special_scopes_formats_and_comments() {
        let mut ic = clipboard();
        ic.formats[0][0].bold = true;
        let grid = vec![vec!["=A9".into()]];
        let targets = vec![(5, 2, 0, 0)];
        let writes = table_paste_writes(&grid, Some(&ic), TablePasteKind::Formats, targets.clone());
        assert!(writes[0].value.is_none());
        assert!(writes[0].format.as_ref().unwrap().bold);
        assert!(writes[0].comment.is_none());
        let writes = table_paste_writes(&grid, Some(&ic), TablePasteKind::All, targets.clone());
        assert!(writes[0].format.is_some());
        assert_eq!(writes[0].comment, Some(None));
        let writes = table_paste_writes(&grid, Some(&ic), TablePasteKind::Values, targets);
        assert_eq!(writes[0].value.as_deref(), Some("9"));
        assert!(writes[0].format.is_none() && writes[0].comment.is_none());
    }
}
