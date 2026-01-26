// Excel file import (xlsx, xls, xlsb, ods) and export (xlsx only)
//
// Import: One-way conversion. Files are converted to VisiGrid's internal model.
// Export: Presentation snapshot for sharing. Not a round-trip format.
//         See docs/features/xlsx-export-spec.md for design rationale.

use std::collections::HashMap;
use std::path::Path;
use std::time::Instant;

use calamine::{open_workbook_auto, Data, Reader, Sheets};
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, FormatUnderline, Workbook as XlsxWorkbook, Worksheet};
use visigrid_engine::cell::{Alignment, BorderStyle, CellFormat, CellValue, DateStyle, NumberFormat, VerticalAlignment};
use visigrid_engine::formula::analyze::tally_unknown_functions;
use visigrid_engine::formula::parser::parse as parse_formula;
use visigrid_engine::sheet::{Sheet, SheetId};
use visigrid_engine::workbook::Workbook;

/// Per-sheet import statistics
#[derive(Debug, Default, Clone)]
pub struct SheetStats {
    pub name: String,
    pub cells_imported: usize,
    pub formulas_imported: usize,
    pub formulas_with_errors: usize,      // Formulas that failed to parse
    pub formulas_with_unknowns: usize,    // Formulas with unknown functions
    pub dates_imported: usize,
    pub times_imported: usize,
    pub truncated_rows: usize,
    pub truncated_cols: usize,
}

/// Result of an Excel import operation
#[derive(Debug, Default)]
pub struct ImportResult {
    /// Per-sheet statistics
    pub sheet_stats: Vec<SheetStats>,
    /// Count of sheets imported
    pub sheets_imported: usize,
    /// Total cells imported
    pub cells_imported: usize,
    /// Total formulas imported
    pub formulas_imported: usize,
    /// Total formulas that failed to parse
    pub formulas_failed: usize,
    /// Total formulas with unknown functions
    pub formulas_with_unknowns: usize,
    /// Total dates/times imported
    pub dates_imported: usize,
    /// Whether any truncation occurred
    pub truncated: bool,
    /// Whether 1904 date system was detected
    pub is_1904_system: bool,
    /// Unknown/unsupported functions encountered (function name -> count)
    pub unsupported_functions: HashMap<String, usize>,
    /// Actionable warnings (not boilerplate)
    pub warnings: Vec<String>,
    /// Total import duration in milliseconds
    pub import_duration_ms: u128,
    /// Total validations imported
    pub validations_imported: usize,
    /// Total validations skipped (unsupported types)
    pub validations_skipped: usize,
}

impl ImportResult {
    /// Returns a summary message suitable for display
    pub fn summary(&self) -> String {
        let mut parts = vec![
            format!("{} sheet{}", self.sheets_imported, if self.sheets_imported == 1 { "" } else { "s" }),
            format!("{} cells", self.cells_imported),
        ];
        if self.formulas_imported > 0 {
            parts.push(format!("{} formulas", self.formulas_imported));
        }
        parts.join(", ")
    }

    /// Returns true if there are actionable warnings
    pub fn has_warnings(&self) -> bool {
        self.truncated || self.formulas_failed > 0 || !self.unsupported_functions.is_empty() || !self.warnings.is_empty()
    }

    /// Returns a single-line warning for status bar (only actionable issues)
    pub fn warning_summary(&self) -> Option<String> {
        let mut issues = Vec::new();

        if self.truncated {
            issues.push("data truncated".to_string());
        }

        if self.formulas_failed > 0 {
            issues.push(format!("{} formula parse errors", self.formulas_failed));
        }

        let unsupported_count: usize = self.unsupported_functions.values().sum();
        if unsupported_count > 0 {
            issues.push(format!("{} unsupported functions", unsupported_count));
        }

        if issues.is_empty() {
            None
        } else {
            Some(format!("Import issues: {}", issues.join(", ")))
        }
    }

    /// Get top N unsupported functions by usage count
    pub fn top_unsupported_functions(&self, n: usize) -> Vec<(&String, &usize)> {
        let mut sorted: Vec<_> = self.unsupported_functions.iter().collect();
        sorted.sort_by(|a, b| b.1.cmp(a.1));
        sorted.truncate(n);
        sorted
    }
}

/// Maximum number of cells to import (prevents DoS from huge files)
const MAX_CELLS: usize = 5_000_000;

/// Maximum dimensions for a sheet
const MAX_ROWS: usize = 65536;
const MAX_COLS: usize = 256;

/// Import an Excel file (xlsx, xls, xlsb, ods)
pub fn import(path: &Path) -> Result<(Workbook, ImportResult), String> {
    let start_time = Instant::now();

    let mut workbook: Sheets<_> = open_workbook_auto(path)
        .map_err(|e| format!("Failed to open Excel file: {}", e))?;

    let mut result = ImportResult::default();
    let mut sheets: Vec<Sheet> = Vec::new();
    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();

    if sheet_names.is_empty() {
        return Err("Excel file contains no sheets".to_string());
    }

    let mut total_cells = 0;
    let mut hit_cell_limit = false;
    let mut next_sheet_id: u64 = 1;

    for sheet_name in &sheet_names {
        let range = workbook.worksheet_range(sheet_name)
            .map_err(|e| format!("Failed to read sheet '{}': {}", sheet_name, e))?;

        let (height, width) = range.get_size();

        let mut stats = SheetStats {
            name: sheet_name.clone(),
            ..Default::default()
        };

        // Skip empty sheets but still create them
        if height == 0 || width == 0 {
            let mut sheet = Sheet::new_with_name(SheetId(next_sheet_id), MAX_ROWS, MAX_COLS, sheet_name);
            next_sheet_id += 1;

            // Import validations even for empty sheets
            let (imported, skipped) = import_validation_rules(path, sheet_name, &mut sheet);
            result.validations_imported += imported;
            result.validations_skipped += skipped;

            sheets.push(sheet);
            result.sheets_imported += 1;
            result.sheet_stats.push(stats);
            continue;
        }

        // Cap dimensions to our limits
        let effective_rows = height.min(MAX_ROWS);
        let effective_cols = width.min(MAX_COLS);

        if height > MAX_ROWS || width > MAX_COLS {
            stats.truncated_rows = height.saturating_sub(MAX_ROWS);
            stats.truncated_cols = width.saturating_sub(MAX_COLS);
            result.truncated = true;
            result.warnings.push(format!(
                "Sheet '{}' truncated from {}x{} to {}x{}",
                sheet_name, height, width, effective_rows, effective_cols
            ));
        }

        let mut sheet = Sheet::new_with_name(SheetId(next_sheet_id), MAX_ROWS, MAX_COLS, sheet_name);
        next_sheet_id += 1;

        for (row_idx, row) in range.rows().enumerate() {
            if row_idx >= effective_rows {
                break;
            }

            for (col_idx, cell) in row.iter().enumerate() {
                if col_idx >= effective_cols {
                    break;
                }

                // Check cell limit
                if total_cells >= MAX_CELLS {
                    if !hit_cell_limit {
                        hit_cell_limit = true;
                        result.truncated = true;
                        result.warnings.push(format!(
                            "Import stopped at {} cells (limit reached)",
                            MAX_CELLS
                        ));
                    }
                    break;
                }

                let target_row = row_idx;
                let target_col = col_idx;

                match cell {
                    Data::Empty => {
                        // Skip empty cells
                    }
                    Data::String(s) => {
                        if !s.is_empty() {
                            sheet.set_value(target_row, target_col, s);
                            stats.cells_imported += 1;
                            total_cells += 1;
                        }
                    }
                    Data::Float(n) => {
                        // Format nicely: integers without decimals
                        let value_str = if n.fract() == 0.0 && n.abs() < 1e15 {
                            format!("{}", *n as i64)
                        } else {
                            format!("{}", n)
                        };
                        sheet.set_value(target_row, target_col, &value_str);
                        stats.cells_imported += 1;
                        total_cells += 1;
                    }
                    Data::Int(n) => {
                        let value_str = format!("{}", n);
                        sheet.set_value(target_row, target_col, &value_str);
                        stats.cells_imported += 1;
                        total_cells += 1;
                    }
                    Data::Bool(b) => {
                        // Store as TRUE/FALSE text which will be recognized by formulas
                        sheet.set_value(target_row, target_col, if *b { "TRUE" } else { "FALSE" });
                        stats.cells_imported += 1;
                        total_cells += 1;
                    }
                    Data::Error(e) => {
                        // Store error as text representation
                        let error_str = format!("#{:?}", e);
                        sheet.set_value(target_row, target_col, &error_str);
                        stats.cells_imported += 1;
                        total_cells += 1;
                    }
                    Data::DateTime(dt) => {
                        // Note: calamine's ExcelDateTime stores is_1904 flag internally
                        // but doesn't expose a getter. For now, we assume 1900 system
                        // (most common). The dates feature could enable as_datetime()
                        // which handles this internally.
                        let serial = dt.as_f64();

                        let value_str = format!("{}", serial);
                        sheet.set_value(target_row, target_col, &value_str);

                        // Determine format based on serial value
                        let has_date = serial.floor() > 0.0;
                        let has_time = serial.fract().abs() > 0.0001; // Small epsilon for float comparison

                        let mut format = sheet.get_format(target_row, target_col);
                        format.number_format = if has_date && has_time {
                            NumberFormat::DateTime
                        } else if has_time {
                            NumberFormat::Time
                        } else {
                            NumberFormat::Date { style: DateStyle::Short }
                        };
                        sheet.set_format(target_row, target_col, format);

                        if has_time {
                            stats.times_imported += 1;
                        } else {
                            stats.dates_imported += 1;
                        }
                        stats.cells_imported += 1;
                        total_cells += 1;
                    }
                    Data::DateTimeIso(s) => {
                        // ISO date string - store as text for now
                        sheet.set_value(target_row, target_col, s);
                        stats.cells_imported += 1;
                        total_cells += 1;
                    }
                    Data::DurationIso(s) => {
                        // ISO duration string - store as text
                        sheet.set_value(target_row, target_col, s);
                        stats.cells_imported += 1;
                        total_cells += 1;
                    }
                }
            }

            if hit_cell_limit {
                break;
            }
        }

        // Try to import formulas if available
        if !hit_cell_limit {
            if let Ok(formula_range) = workbook.worksheet_formula(sheet_name) {
                for (row_idx, row) in formula_range.rows().enumerate() {
                    if row_idx >= effective_rows {
                        break;
                    }

                    for (col_idx, formula) in row.iter().enumerate() {
                        if col_idx >= effective_cols {
                            break;
                        }

                        // Enforce MAX_CELLS for formulas too
                        if total_cells >= MAX_CELLS {
                            if !hit_cell_limit {
                                hit_cell_limit = true;
                                result.truncated = true;
                                result.warnings.push(format!(
                                    "Formula import stopped at {} cells (limit reached)",
                                    MAX_CELLS
                                ));
                            }
                            break;
                        }

                        if !formula.is_empty() {
                            let target_row = row_idx;
                            let target_col = col_idx;

                            // Check if this cell was empty before (formula adds a new cell)
                            let was_empty = sheet.get_raw(target_row, target_col).is_empty();

                            // Preserve existing format if cell was already set
                            let existing_format = sheet.get_format(target_row, target_col);

                            // Import formula (prepend = if not present)
                            let formula_str = if formula.starts_with('=') {
                                formula.clone()
                            } else {
                                format!("={}", formula)
                            };

                            // Analyze formula for unknown functions
                            match parse_formula(&formula_str) {
                                Ok(ast) => {
                                    // Count unknown functions in this formula
                                    let prev_count: usize = result.unsupported_functions.values().sum();
                                    tally_unknown_functions(&ast, &mut result.unsupported_functions);
                                    let new_count: usize = result.unsupported_functions.values().sum();

                                    // If any new unknown functions were found, track the cell
                                    if new_count > prev_count {
                                        stats.formulas_with_unknowns += 1;
                                    }
                                }
                                Err(_) => {
                                    // Formula failed to parse - track it
                                    stats.formulas_with_errors += 1;
                                }
                            }

                            sheet.set_value(target_row, target_col, &formula_str);
                            sheet.set_format(target_row, target_col, existing_format);

                            stats.formulas_imported += 1;

                            // If cell was empty, this formula adds a new cell
                            if was_empty {
                                stats.cells_imported += 1;
                                total_cells += 1;
                            }
                        }
                    }

                    if hit_cell_limit {
                        break;
                    }
                }
            }
        }

        // Update totals
        result.cells_imported += stats.cells_imported;
        result.formulas_imported += stats.formulas_imported;
        result.formulas_failed += stats.formulas_with_errors;
        result.formulas_with_unknowns += stats.formulas_with_unknowns;
        result.dates_imported += stats.dates_imported + stats.times_imported;

        // Import validation rules from worksheet XML
        let (imported, skipped) = import_validation_rules(path, sheet_name, &mut sheet);
        result.validations_imported += imported;
        result.validations_skipped += skipped;

        sheets.push(sheet);
        result.sheets_imported += 1;
        result.sheet_stats.push(stats);
    }

    if sheets.is_empty() {
        return Err("No sheets could be imported".to_string());
    }

    // Record import duration
    result.import_duration_ms = start_time.elapsed().as_millis();

    let mut workbook = Workbook::from_sheets(sheets, 0);

    // Rebuild dependency graph after loading all data
    workbook.rebuild_dep_graph();

    Ok((workbook, result))
}

// =============================================================================
// XLSX Export
// =============================================================================

/// A formula that was converted to a value during export
#[derive(Debug, Clone)]
pub struct ConvertedFormula {
    /// Sheet name
    pub sheet: String,
    /// Cell address (e.g., "B5")
    pub address: String,
    /// Original formula (e.g., "=CUSTOM_FUNC(A1)")
    pub formula: String,
    /// The value it was converted to
    pub value: String,
}

/// A large number that was exported as text to preserve precision
#[derive(Debug, Clone)]
pub struct PrecisionWarning {
    /// Sheet name
    pub sheet: String,
    /// Cell address (e.g., "B5")
    pub address: String,
    /// The number (as string, since it's too large for f64 precision)
    pub value: String,
}

/// Excel has 15-digit precision. Numbers with more significant digits lose precision.
#[allow(dead_code)]
const EXCEL_MAX_SAFE_DIGITS: usize = 15;

/// Check if a number exceeds Excel's safe precision (15 significant digits)
/// Returns true if the integer part has more than 15 digits
fn exceeds_excel_precision(n: f64) -> bool {
    if !n.is_finite() {
        return false;
    }
    // Check if integer part exceeds 15 digits
    let abs_int = n.trunc().abs();
    if abs_int >= 1e15 {
        return true;
    }
    false
}

/// Result of an Excel export operation
#[derive(Debug, Default)]
pub struct ExportResult {
    /// Number of sheets exported
    pub sheets_exported: usize,
    /// Total cells exported
    pub cells_exported: usize,
    /// Formulas exported as formulas (not converted to values)
    pub formulas_exported: usize,
    /// Formulas that were converted to values (couldn't be expressed in Excel)
    pub formulas_as_values: usize,
    /// Detailed list of converted formulas (for user review)
    pub converted_formulas: Vec<ConvertedFormula>,
    /// Numbers exported as text due to precision limits (>15 digits)
    pub precision_warnings: Vec<PrecisionWarning>,
    /// Validation rules exported
    pub validations_exported: usize,
    /// Validation rules skipped (unsupported types)
    pub validations_skipped: usize,
    /// Export duration in milliseconds
    pub export_duration_ms: u128,
    /// Warnings generated during export
    pub warnings: Vec<String>,
}

impl ExportResult {
    /// Returns a summary message suitable for display
    pub fn summary(&self) -> String {
        let mut parts = vec![
            format!("{} sheet{}", self.sheets_exported, if self.sheets_exported == 1 { "" } else { "s" }),
            format!("{} cells", self.cells_exported),
        ];
        if self.formulas_exported > 0 {
            parts.push(format!("{} formulas", self.formulas_exported));
        }
        parts.join(", ")
    }

    /// Returns true if there are warnings
    pub fn has_warnings(&self) -> bool {
        self.formulas_as_values > 0 || !self.precision_warnings.is_empty() || !self.warnings.is_empty()
    }

    /// Returns a single-line warning for status bar
    pub fn warning_summary(&self) -> Option<String> {
        let mut parts = Vec::new();

        if self.formulas_as_values > 0 {
            parts.push(format!("{} formulas as values", self.formulas_as_values));
        }
        if !self.precision_warnings.is_empty() {
            parts.push(format!("{} numbers as text", self.precision_warnings.len()));
        }

        if !parts.is_empty() {
            Some(parts.join(", "))
        } else if !self.warnings.is_empty() {
            Some(self.warnings.join(", "))
        } else {
            None
        }
    }

    /// Returns a detailed report of converted formulas for user review
    /// Format: one line per formula, e.g., "Sheet1!B5: =CUSTOM_FUNC(A1) → #ERR"
    pub fn converted_formulas_report(&self) -> String {
        if self.converted_formulas.is_empty() {
            return String::new();
        }

        let mut lines = Vec::new();
        lines.push(format!(
            "{} formulas could not be exported and were converted to values:",
            self.converted_formulas.len()
        ));

        for cf in &self.converted_formulas {
            lines.push(format!("  {}!{}: {} → {}", cf.sheet, cf.address, cf.formula, cf.value));
        }

        lines.join("\n")
    }

    /// Returns a detailed report of precision warnings for user review
    pub fn precision_warnings_report(&self) -> String {
        if self.precision_warnings.is_empty() {
            return String::new();
        }

        let mut lines = Vec::new();
        lines.push(format!(
            "{} numbers exceeded Excel's 15-digit precision limit.",
            self.precision_warnings.len()
        ));
        lines.push("These were exported as text to preserve exact values (they won't work in Excel formulas):".to_string());

        for pw in &self.precision_warnings {
            lines.push(format!("  {}!{}: {}", pw.sheet, pw.address, pw.value));
        }

        lines.join("\n")
    }

    /// Returns a full export report with file context - suitable for copy/paste sharing
    pub fn full_report_with_context(&self, filename: &str) -> String {
        let mut lines = Vec::new();

        // Header
        lines.push("VisiGrid Export Report".to_string());
        lines.push("=".repeat(50));
        lines.push(String::new());

        // File info
        lines.push(format!("File: {}", filename));
        lines.push(format!("Exported: {} sheets, {} cells, {} formulas",
            self.sheets_exported, self.cells_exported, self.formulas_exported));
        lines.push(format!("Duration: {} ms", self.export_duration_ms));

        // Warnings summary
        if self.has_warnings() {
            lines.push(String::new());
            lines.push(format!("Warnings: {} formulas converted to values, {} numbers as text",
                self.formulas_as_values, self.precision_warnings.len()));
        }

        lines.push(String::new());
        lines.push("-".repeat(50));

        // Detailed sections
        let formulas_report = self.converted_formulas_report();
        if !formulas_report.is_empty() {
            lines.push(String::new());
            lines.push(formulas_report);
        }

        let precision_report = self.precision_warnings_report();
        if !precision_report.is_empty() {
            lines.push(String::new());
            lines.push(precision_report);
        }

        lines.join("\n")
    }

    /// Returns a full export report combining all warnings (without file context)
    pub fn full_report(&self) -> String {
        let mut sections = Vec::new();

        let formulas_report = self.converted_formulas_report();
        if !formulas_report.is_empty() {
            sections.push(formulas_report);
        }

        let precision_report = self.precision_warnings_report();
        if !precision_report.is_empty() {
            sections.push(precision_report);
        }

        sections.join("\n\n")
    }
}

/// Layout information for export (column widths, row heights, frozen panes)
#[derive(Debug, Default, Clone)]
pub struct ExportLayout {
    /// Column widths in pixels, keyed by column index
    pub col_widths: HashMap<usize, f32>,
    /// Row heights in pixels, keyed by row index
    pub row_heights: HashMap<usize, f32>,
    /// Number of frozen rows (for freeze panes)
    pub frozen_rows: usize,
    /// Number of frozen columns (for freeze panes)
    pub frozen_cols: usize,
}

/// Convert pixel width to Excel column width (approximate)
/// Excel measures column width in characters (based on default font)
fn pixels_to_excel_width(px: f32) -> f64 {
    (px / 7.0) as f64 // ~7 pixels per character
}

/// Convert pixel height to Excel row height
/// Excel measures row height in points (1/72 inch)
fn pixels_to_excel_height(px: f32) -> f64 {
    (px * 0.75) as f64 // 1 point ≈ 1.33 pixels
}

/// Export a VisiGrid workbook to XLSX format
///
/// # Arguments
/// * `workbook` - The VisiGrid workbook to export
/// * `path` - Path to write the XLSX file
/// * `layouts` - Optional per-sheet layout information (column widths, row heights, frozen panes)
///
/// # Returns
/// * `Ok(ExportResult)` - Export statistics
/// * `Err(String)` - Error message if export failed
pub fn export(
    workbook: &Workbook,
    path: &Path,
    layouts: Option<&[ExportLayout]>,
) -> Result<ExportResult, String> {
    let start_time = Instant::now();
    let mut result = ExportResult::default();

    let mut xlsx_workbook = XlsxWorkbook::new();

    for (sheet_idx, sheet) in workbook.sheets().iter().enumerate() {
        let worksheet = xlsx_workbook
            .add_worksheet()
            .set_name(&sheet.name)
            .map_err(|e| format!("Failed to create sheet '{}': {}", sheet.name, e))?;

        // Get layout for this sheet if provided
        let layout = layouts.and_then(|l| l.get(sheet_idx));

        // Export cells
        let (cells, formulas, as_values, converted, precision) = export_sheet_cells(sheet, worksheet)?;
        result.cells_exported += cells;
        result.formulas_exported += formulas;
        result.formulas_as_values += as_values;
        result.converted_formulas.extend(converted);
        result.precision_warnings.extend(precision);

        // Apply layout (column widths, row heights, frozen panes)
        if let Some(layout) = layout {
            apply_layout(worksheet, layout)?;
        }

        // Export validation rules
        let (exported, skipped) = export_validation_rules(worksheet, sheet)?;
        result.validations_exported += exported;
        result.validations_skipped += skipped;

        result.sheets_exported += 1;
    }

    // Set active sheet
    if let Ok(ws) = xlsx_workbook.worksheet_from_index(workbook.active_sheet_index()) {
        let _ = ws.set_active(true);
    }

    // Save to file
    xlsx_workbook
        .save(path)
        .map_err(|e| format!("Failed to save XLSX file: {}", e))?;

    result.export_duration_ms = start_time.elapsed().as_millis();
    Ok(result)
}

/// Convert column index to Excel column letter (0 = A, 25 = Z, 26 = AA, etc.)
fn col_to_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;
    loop {
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    result
}

/// Convert row/col to Excel cell address (e.g., "A1", "B5", "AA100")
fn cell_address(row: usize, col: usize) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

/// Export cells from a VisiGrid sheet to an Excel worksheet
/// Returns (cells_exported, formulas_exported, formulas_as_values, converted_formulas, precision_warnings)
fn export_sheet_cells(
    sheet: &Sheet,
    worksheet: &mut Worksheet,
) -> Result<(usize, usize, usize, Vec<ConvertedFormula>, Vec<PrecisionWarning>), String> {
    let mut cells_exported = 0;
    let mut formulas_exported = 0;
    let mut formulas_as_values = 0;
    let mut converted_formulas = Vec::new();
    let mut precision_warnings = Vec::new();

    // Iterate over all cells in the sheet
    for ((row, col), cell) in sheet.cells_iter() {
        // Skip spill receiver cells - they'll be filled by Excel when recalculating
        if cell.is_spill_receiver() {
            continue;
        }

        let row32 = *row as u32;
        let col16 = *col as u16;

        // Build format for this cell
        let format = build_excel_format(&cell.format);

        match &cell.value {
            CellValue::Empty => {
                // Only write format if cell has formatting
                if has_formatting(&cell.format) {
                    worksheet
                        .write_blank(row32, col16, &format)
                        .map_err(|e| format!("Failed to write cell ({}, {}): {}", row, col, e))?;
                    cells_exported += 1;
                }
            }
            CellValue::Text(s) => {
                worksheet
                    .write_string_with_format(row32, col16, s, &format)
                    .map_err(|e| format!("Failed to write cell ({}, {}): {}", row, col, e))?;
                cells_exported += 1;
            }
            CellValue::Number(n) => {
                // Check for precision loss (>15 significant digits)
                if exceeds_excel_precision(*n) {
                    // Export as text to preserve exact value
                    let text_value = format!("{}", *n as i64); // Format as integer string
                    worksheet
                        .write_string_with_format(row32, col16, &text_value, &format)
                        .map_err(|e| format!("Failed to write cell ({}, {}): {}", row, col, e))?;

                    precision_warnings.push(PrecisionWarning {
                        sheet: sheet.name.clone(),
                        address: cell_address(*row, *col),
                        value: text_value,
                    });
                } else {
                    // Safe to export as number
                    let format = apply_number_format(format, &cell.format.number_format);
                    worksheet
                        .write_number_with_format(row32, col16, *n, &format)
                        .map_err(|e| format!("Failed to write cell ({}, {}): {}", row, col, e))?;
                }
                cells_exported += 1;
            }
            CellValue::Formula { source, ast } => {
                // Try to export as formula if it has a valid AST
                if ast.is_some() {
                    // Export the formula string (strip leading '=')
                    let formula_str = source.strip_prefix('=').unwrap_or(source);
                    let format = apply_number_format(format, &cell.format.number_format);

                    worksheet
                        .write_formula_with_format(row32, col16, formula_str, &format)
                        .map_err(|e| format!("Failed to write formula ({}, {}): {}", row, col, e))?;
                    formulas_exported += 1;
                } else {
                    // Invalid formula - export computed value instead
                    let display = sheet.get_formatted_display(*row, *col);
                    if let Ok(n) = display.parse::<f64>() {
                        let format = apply_number_format(format, &cell.format.number_format);
                        worksheet
                            .write_number_with_format(row32, col16, n, &format)
                            .map_err(|e| format!("Failed to write cell ({}, {}): {}", row, col, e))?;
                    } else {
                        worksheet
                            .write_string_with_format(row32, col16, &display, &format)
                            .map_err(|e| format!("Failed to write cell ({}, {}): {}", row, col, e))?;
                    }

                    // Track this conversion for user review
                    converted_formulas.push(ConvertedFormula {
                        sheet: sheet.name.clone(),
                        address: cell_address(*row, *col),
                        formula: source.clone(),
                        value: display,
                    });

                    formulas_as_values += 1;
                }
                cells_exported += 1;
            }
        }
    }

    Ok((cells_exported, formulas_exported, formulas_as_values, converted_formulas, precision_warnings))
}

/// Export validation rules for a sheet
///
/// Returns (exported_count, skipped_count).
/// Skipped rules are those with unsupported types (Date, Time, TextLength, Custom).
fn export_validation_rules(
    worksheet: &mut Worksheet,
    sheet: &Sheet,
) -> Result<(usize, usize), String> {
    use crate::xlsx_validation::rule_to_xlsx;

    let mut exported = 0;
    let mut skipped = 0;

    for (range, rule) in sheet.validations.iter() {
        match rule_to_xlsx(rule) {
            Some(dv) => {
                // rust_xlsxwriter uses 0-based row/col as u32/u16
                worksheet
                    .add_data_validation(
                        range.start_row as u32,
                        range.start_col as u16,
                        range.end_row as u32,
                        range.end_col as u16,
                        &dv,
                    )
                    .map_err(|e| format!("Failed to add validation: {}", e))?;
                exported += 1;
            }
            None => {
                // Unsupported validation type (Date, Time, TextLength, Custom)
                skipped += 1;
            }
        }
    }

    Ok((exported, skipped))
}

/// Import validation rules for a sheet from XLSX
///
/// Returns (imported_count, skipped_count).
/// Skipped rules are those with unsupported types (Date, Time, TextLength, Custom).
fn import_validation_rules(
    xlsx_path: &Path,
    sheet_name: &str,
    sheet: &mut Sheet,
) -> (usize, usize) {
    use crate::xlsx_validation::parse_sheet_validations;

    match parse_sheet_validations(xlsx_path, sheet_name) {
        Ok(validations) => {
            let mut imported = 0;
            for v in validations {
                sheet.validations.set(v.range, v.rule);
                imported += 1;
            }
            (imported, 0) // Skipping is handled in parse_sheet_validations
        }
        Err(_) => {
            // Validation parsing failed - not fatal, just skip
            // This can happen if the sheet has no validations or XML structure differs
            (0, 0)
        }
    }
}

/// Build an Excel Format from VisiGrid CellFormat
fn build_excel_format(cell_format: &CellFormat) -> Format {
    let mut format = Format::new();

    // Font styling
    if cell_format.bold {
        format = format.set_bold();
    }
    if cell_format.italic {
        format = format.set_italic();
    }
    if cell_format.underline {
        format = format.set_underline(FormatUnderline::Single);
    }
    if cell_format.strikethrough {
        format = format.set_font_strikethrough();
    }

    // Horizontal alignment
    format = match cell_format.alignment {
        Alignment::General => format, // Excel default: numbers right, text left
        Alignment::Left => format.set_align(FormatAlign::Left),
        Alignment::Center => format.set_align(FormatAlign::Center),
        Alignment::Right => format.set_align(FormatAlign::Right),
    };

    // Vertical alignment
    format = match cell_format.vertical_alignment {
        VerticalAlignment::Top => format.set_align(FormatAlign::Top),
        VerticalAlignment::Middle => format.set_align(FormatAlign::VerticalCenter),
        VerticalAlignment::Bottom => format.set_align(FormatAlign::Bottom),
    };

    // Text wrap (from TextOverflow::Wrap)
    if cell_format.text_overflow == visigrid_engine::cell::TextOverflow::Wrap {
        format = format.set_text_wrap();
    }

    // Background fill color
    if let Some([r, g, b, _]) = cell_format.background_color {
        // Excel uses ARGB, but rust_xlsxwriter uses RGB hex
        let color = rust_xlsxwriter::Color::RGB(
            ((r as u32) << 16) | ((g as u32) << 8) | (b as u32)
        );
        format = format.set_background_color(color);
    }

    // Cell borders
    if cell_format.border_top.style != BorderStyle::None {
        format = format.set_border_top(border_style_to_xlsx(cell_format.border_top.style));
    }
    if cell_format.border_right.style != BorderStyle::None {
        format = format.set_border_right(border_style_to_xlsx(cell_format.border_right.style));
    }
    if cell_format.border_bottom.style != BorderStyle::None {
        format = format.set_border_bottom(border_style_to_xlsx(cell_format.border_bottom.style));
    }
    if cell_format.border_left.style != BorderStyle::None {
        format = format.set_border_left(border_style_to_xlsx(cell_format.border_left.style));
    }

    format
}

/// Convert VisiGrid BorderStyle to rust_xlsxwriter FormatBorder
fn border_style_to_xlsx(style: BorderStyle) -> FormatBorder {
    match style {
        BorderStyle::None => FormatBorder::None,
        BorderStyle::Thin => FormatBorder::Thin,
        BorderStyle::Medium => FormatBorder::Medium,
        BorderStyle::Thick => FormatBorder::Thick,
    }
}

/// Apply number format to an Excel Format
fn apply_number_format(format: Format, number_format: &NumberFormat) -> Format {
    match number_format {
        NumberFormat::General => format,
        NumberFormat::Number { decimals } => {
            let pattern = if *decimals == 0 {
                "0".to_string()
            } else {
                format!("0.{}", "0".repeat(*decimals as usize))
            };
            format.set_num_format(&pattern)
        }
        NumberFormat::Currency { decimals } => {
            let pattern = if *decimals == 0 {
                "$#,##0".to_string()
            } else {
                format!("$#,##0.{}", "0".repeat(*decimals as usize))
            };
            format.set_num_format(&pattern)
        }
        NumberFormat::Percent { decimals } => {
            let pattern = if *decimals == 0 {
                "0%".to_string()
            } else {
                format!("0.{}%", "0".repeat(*decimals as usize))
            };
            format.set_num_format(&pattern)
        }
        NumberFormat::Date { style } => {
            let pattern = match style {
                DateStyle::Short => "m/d/yyyy",
                DateStyle::Long => "mmmm d, yyyy",
                DateStyle::Iso => "yyyy-mm-dd",
            };
            format.set_num_format(pattern)
        }
        NumberFormat::Time => format.set_num_format("h:mm:ss"),
        NumberFormat::DateTime => format.set_num_format("m/d/yyyy h:mm:ss"),
    }
}

/// Check if a CellFormat has any non-default formatting
fn has_formatting(format: &CellFormat) -> bool {
    format.bold
        || format.italic
        || format.underline
        || format.strikethrough
        || format.alignment != Alignment::General
        || format.vertical_alignment != VerticalAlignment::Middle
        || format.number_format != NumberFormat::General
        || format.font_family.is_some()
        || format.background_color.is_some()
        || format.border_top.style != BorderStyle::None
        || format.border_right.style != BorderStyle::None
        || format.border_bottom.style != BorderStyle::None
        || format.border_left.style != BorderStyle::None
}

/// Apply layout settings (column widths, row heights, frozen panes) to worksheet
fn apply_layout(worksheet: &mut Worksheet, layout: &ExportLayout) -> Result<(), String> {
    // Apply column widths
    for (col, width_px) in &layout.col_widths {
        let excel_width = pixels_to_excel_width(*width_px);
        worksheet
            .set_column_width(*col as u16, excel_width)
            .map_err(|e| format!("Failed to set column {} width: {}", col, e))?;
    }

    // Apply row heights
    for (row, height_px) in &layout.row_heights {
        let excel_height = pixels_to_excel_height(*height_px);
        worksheet
            .set_row_height(*row as u32, excel_height)
            .map_err(|e| format!("Failed to set row {} height: {}", row, e))?;
    }

    // Apply frozen panes
    if layout.frozen_rows > 0 || layout.frozen_cols > 0 {
        worksheet.set_freeze_panes(layout.frozen_rows as u32, layout.frozen_cols as u16)
            .map_err(|e| format!("Failed to set freeze panes: {}", e))?;
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_import_result_summary() {
        let mut result = ImportResult::default();
        result.sheets_imported = 1;
        result.cells_imported = 100;
        result.formulas_imported = 0;

        assert_eq!(result.summary(), "1 sheet, 100 cells");

        result.sheets_imported = 3;
        result.formulas_imported = 25;
        assert_eq!(result.summary(), "3 sheets, 100 cells, 25 formulas");
    }

    #[test]
    fn test_import_result_warnings_only_actionable() {
        // No warnings if nothing went wrong
        let result = ImportResult::default();
        assert!(!result.has_warnings());
        assert!(result.warning_summary().is_none());

        // Warning if truncated
        let mut result = ImportResult::default();
        result.truncated = true;
        assert!(result.has_warnings());
        assert!(result.warning_summary().unwrap().contains("truncated"));

        // Warning if unsupported functions
        let mut result = ImportResult::default();
        result.unsupported_functions.insert("XLOOKUP".to_string(), 5);
        assert!(result.has_warnings());
        assert!(result.warning_summary().unwrap().contains("unsupported"));
    }

    #[test]
    fn test_top_unsupported_functions() {
        let mut result = ImportResult::default();
        result.unsupported_functions.insert("XLOOKUP".to_string(), 10);
        result.unsupported_functions.insert("TEXTJOIN".to_string(), 5);
        result.unsupported_functions.insert("LET".to_string(), 3);

        let top = result.top_unsupported_functions(2);
        assert_eq!(top.len(), 2);
        assert_eq!(top[0].0, "XLOOKUP");
        assert_eq!(*top[0].1, 10);
        assert_eq!(top[1].0, "TEXTJOIN");
        assert_eq!(*top[1].1, 5);
    }

    // ========================================================================
    // Export tests
    // ========================================================================

    #[test]
    fn test_export_result_summary() {
        let mut result = ExportResult::default();
        result.sheets_exported = 1;
        result.cells_exported = 100;
        result.formulas_exported = 0;

        assert_eq!(result.summary(), "1 sheet, 100 cells");

        result.sheets_exported = 3;
        result.formulas_exported = 25;
        assert_eq!(result.summary(), "3 sheets, 100 cells, 25 formulas");
    }

    #[test]
    fn test_export_result_warnings() {
        // No warnings if nothing converted
        let result = ExportResult::default();
        assert!(!result.has_warnings());
        assert!(result.warning_summary().is_none());

        // Warning if formulas converted to values
        let mut result = ExportResult::default();
        result.formulas_as_values = 5;
        assert!(result.has_warnings());
        assert!(result.warning_summary().unwrap().contains("5 formulas"));
    }

    #[test]
    fn test_pixels_to_excel_width() {
        // Test conversion - 7 pixels per character (approximate)
        assert!((pixels_to_excel_width(70.0) - 10.0).abs() < 0.01);
        assert!((pixels_to_excel_width(7.0) - 1.0).abs() < 0.01);
    }

    #[test]
    fn test_pixels_to_excel_height() {
        // Test conversion - 1 point = ~1.33 pixels
        assert!((pixels_to_excel_height(100.0) - 75.0).abs() < 0.01);
    }

    #[test]
    fn test_export_basic() {
        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Hello");
        workbook.active_sheet_mut().set_value(0, 1, "123");
        workbook.active_sheet_mut().set_value(1, 0, "=A1");

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_export.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();

        assert_eq!(result.sheets_exported, 1);
        assert_eq!(result.cells_exported, 3);
        assert_eq!(result.formulas_exported, 1);
        assert!(export_path.exists());

        // Verify file has content (XLSX is a ZIP, should have meaningful size)
        let metadata = std::fs::metadata(&export_path).unwrap();
        assert!(metadata.len() > 100); // XLSX files have significant overhead
    }

    #[test]
    fn test_export_with_formatting() {
        use visigrid_engine::cell::{CellFormat, NumberFormat};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Set a cell with currency format
        sheet.set_value(0, 0, "1234.56");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Currency { decimals: 2 };
        format.bold = true;
        sheet.set_format(0, 0, format);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_format.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 1);
        assert!(export_path.exists());
    }

    #[test]
    fn test_export_with_layout() {
        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Test");

        let mut layout = ExportLayout::default();
        layout.col_widths.insert(0, 140.0); // Wide column
        layout.row_heights.insert(0, 40.0);  // Tall row
        layout.frozen_rows = 1;

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_layout.xlsx");

        let result = export(&workbook, &export_path, Some(&[layout])).unwrap();
        assert_eq!(result.sheets_exported, 1);
        assert!(export_path.exists());
    }

    #[test]
    fn test_export_multiple_sheets() {
        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Sheet1 Data");
        workbook.add_sheet();
        workbook.sheet_mut(1).unwrap().set_value(0, 0, "Sheet2 Data");

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_multi.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.sheets_exported, 2);
        assert_eq!(result.cells_exported, 2);
    }

    #[test]
    fn test_export_with_borders() {
        use visigrid_engine::cell::{CellBorder, CellFormat};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Create a 5x5 bordered table
        let thin = CellBorder::thin();
        for row in 0..5 {
            for col in 0..5 {
                sheet.set_value(row, col, &format!("R{}C{}", row, col));
                let mut format = CellFormat::default();
                format.border_top = thin;
                format.border_right = thin;
                format.border_bottom = thin;
                format.border_left = thin;
                sheet.set_format(row, col, format);
            }
        }

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_borders.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 25);
        assert!(export_path.exists());

        // File should have meaningful size (borders add to XLSX complexity)
        let metadata = std::fs::metadata(&export_path).unwrap();
        assert!(metadata.len() > 100);
    }

    #[test]
    fn test_has_formatting() {
        use visigrid_engine::cell::CellFormat;

        // Default format has no formatting
        let default_format = CellFormat::default();
        assert!(!has_formatting(&default_format));

        // Bold has formatting
        let mut bold_format = CellFormat::default();
        bold_format.bold = true;
        assert!(has_formatting(&bold_format));

        // Center alignment has formatting
        let mut center_format = CellFormat::default();
        center_format.alignment = Alignment::Center;
        assert!(has_formatting(&center_format));

        // Border has formatting
        use visigrid_engine::cell::CellBorder;
        let mut border_format = CellFormat::default();
        border_format.border_top = CellBorder::thin();
        assert!(has_formatting(&border_format));
    }

    // ========================================================================
    // Validation export
    // ========================================================================

    #[test]
    fn test_export_with_validation_list() {
        use visigrid_engine::validation::{CellRange, ValidationRule};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Set up some data
        sheet.set_value(0, 0, "Status");
        sheet.set_value(1, 0, "Active");

        // Add list validation to A2:A10
        let rule = ValidationRule::list_inline(vec![
            "Active".into(),
            "Inactive".into(),
            "Pending".into(),
        ]);
        let range = CellRange::new(1, 0, 9, 0);
        sheet.validations.set(range, rule);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_validation.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();

        assert_eq!(result.validations_exported, 1);
        assert_eq!(result.validations_skipped, 0);
        assert!(export_path.exists());
    }

    #[test]
    fn test_export_with_validation_numeric() {
        use visigrid_engine::validation::{CellRange, NumericConstraint, ValidationRule};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        sheet.set_value(0, 0, "Value");
        sheet.set_value(1, 0, "50");

        // Add whole number validation (1-100)
        let rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
        let range = CellRange::new(1, 0, 9, 0);
        sheet.validations.set(range, rule);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_numeric_validation.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();

        assert_eq!(result.validations_exported, 1);
        assert_eq!(result.validations_skipped, 0);
    }

    #[test]
    fn test_export_with_unsupported_validation() {
        use visigrid_engine::validation::{CellRange, NumericConstraint, ValidationRule, ValidationType};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        sheet.set_value(0, 0, "Text");
        sheet.set_value(1, 0, "Hello");

        // Add text length validation (not yet supported)
        let rule = ValidationRule::new(ValidationType::TextLength(NumericConstraint::between(1, 50)));
        let range = CellRange::new(1, 0, 9, 0);
        sheet.validations.set(range, rule);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_unsupported.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();

        // TextLength is skipped in Phase 5A
        assert_eq!(result.validations_exported, 0);
        assert_eq!(result.validations_skipped, 1);
    }

    #[test]
    fn test_export_validation_count_summary() {
        use visigrid_engine::validation::{CellRange, NumericConstraint, ValidationRule, ValidationType};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Add multiple validations: some supported, some not
        let list_rule = ValidationRule::list_inline(vec!["A".into(), "B".into()]);
        sheet.validations.set(CellRange::new(0, 0, 4, 0), list_rule);

        let whole_rule = ValidationRule::whole_number(NumericConstraint::between(0, 100));
        sheet.validations.set(CellRange::new(0, 1, 4, 1), whole_rule);

        let decimal_rule = ValidationRule::decimal(NumericConstraint::greater_than(0.0));
        sheet.validations.set(CellRange::new(0, 2, 4, 2), decimal_rule);

        // Unsupported: Date, Time, TextLength, Custom
        let date_rule = ValidationRule::new(ValidationType::Date(NumericConstraint::between(0, 100)));
        sheet.validations.set(CellRange::new(0, 3, 4, 3), date_rule);

        let custom_rule = ValidationRule::custom("=A1>0");
        sheet.validations.set(CellRange::new(0, 4, 4, 4), custom_rule);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_mixed_validation.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();

        // 3 supported (List, WholeNumber, Decimal), 2 skipped (Date, Custom)
        assert_eq!(result.validations_exported, 3);
        assert_eq!(result.validations_skipped, 2);
    }

    // ========================================================================
    // Validation round-trip tests
    // ========================================================================

    #[test]
    fn test_validation_roundtrip_list() {
        use visigrid_engine::validation::{CellRange, ListSource, ValidationRule, ValidationType};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Create list validation
        let mut rule = ValidationRule::list_inline(vec!["Open".into(), "In Progress".into(), "Closed".into()]);
        rule.ignore_blank = true;
        rule.show_dropdown = true;
        sheet.validations.set(CellRange::new(1, 1, 99, 1), rule);

        // Export
        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("roundtrip_list.xlsx");
        let export_result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(export_result.validations_exported, 1);

        // Import
        let (imported_workbook, import_result) = import(&export_path).unwrap();
        assert_eq!(import_result.validations_imported, 1);

        // Verify the rule was preserved
        let imported_sheet = imported_workbook.active_sheet();
        let imported_rule = imported_sheet.validations.get(50, 1);
        assert!(imported_rule.is_some(), "Validation rule should exist");

        let imported_rule = imported_rule.unwrap();
        assert!(imported_rule.ignore_blank);
        assert!(imported_rule.show_dropdown);

        match &imported_rule.rule_type {
            ValidationType::List(ListSource::Inline(items)) => {
                assert_eq!(items, &vec!["Open", "In Progress", "Closed"]);
            }
            _ => panic!("Expected inline list validation"),
        }
    }

    #[test]
    fn test_validation_roundtrip_whole_number() {
        use visigrid_engine::validation::{CellRange, ComparisonOperator, ConstraintValue, NumericConstraint, ValidationRule, ValidationType};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Create whole number validation (between 1 and 100)
        let mut rule = ValidationRule::whole_number(NumericConstraint::between(1, 100));
        rule.ignore_blank = false;
        sheet.validations.set(CellRange::new(0, 2, 49, 2), rule);

        // Export
        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("roundtrip_whole.xlsx");
        let export_result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(export_result.validations_exported, 1);

        // Import
        let (imported_workbook, import_result) = import(&export_path).unwrap();
        assert_eq!(import_result.validations_imported, 1);

        // Verify the rule was preserved
        let imported_sheet = imported_workbook.active_sheet();
        let imported_rule = imported_sheet.validations.get(25, 2);
        assert!(imported_rule.is_some(), "Validation rule should exist");

        let imported_rule = imported_rule.unwrap();
        assert!(!imported_rule.ignore_blank);

        match &imported_rule.rule_type {
            ValidationType::WholeNumber(constraint) => {
                assert!(matches!(constraint.operator, ComparisonOperator::Between));
                assert!(matches!(constraint.value1, ConstraintValue::Number(n) if (n - 1.0).abs() < 0.001));
                assert!(matches!(constraint.value2, Some(ConstraintValue::Number(n)) if (n - 100.0).abs() < 0.001));
            }
            _ => panic!("Expected whole number validation"),
        }
    }

    #[test]
    fn test_validation_roundtrip_decimal() {
        use visigrid_engine::validation::{CellRange, ComparisonOperator, ConstraintValue, NumericConstraint, ValidationRule, ValidationType};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Create decimal validation (greater than 0)
        let rule = ValidationRule::decimal(NumericConstraint::greater_than(0.0));
        sheet.validations.set(CellRange::new(0, 3, 19, 3), rule);

        // Export
        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("roundtrip_decimal.xlsx");
        let export_result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(export_result.validations_exported, 1);

        // Import
        let (imported_workbook, import_result) = import(&export_path).unwrap();
        assert_eq!(import_result.validations_imported, 1);

        // Verify the rule was preserved
        let imported_sheet = imported_workbook.active_sheet();
        let imported_rule = imported_sheet.validations.get(10, 3);
        assert!(imported_rule.is_some(), "Validation rule should exist");

        let imported_rule = imported_rule.unwrap();

        match &imported_rule.rule_type {
            ValidationType::Decimal(constraint) => {
                assert!(matches!(constraint.operator, ComparisonOperator::GreaterThan));
                assert!(matches!(constraint.value1, ConstraintValue::Number(n) if n.abs() < 0.001));
            }
            _ => panic!("Expected decimal validation"),
        }
    }

    #[test]
    fn test_validation_roundtrip_multiple() {
        use visigrid_engine::validation::{CellRange, NumericConstraint, ValidationRule, ValidationType};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Create multiple validations
        let list_rule = ValidationRule::list_inline(vec!["A".into(), "B".into(), "C".into()]);
        sheet.validations.set(CellRange::new(0, 0, 9, 0), list_rule);

        let whole_rule = ValidationRule::whole_number(NumericConstraint::between(0, 50));
        sheet.validations.set(CellRange::new(0, 1, 9, 1), whole_rule);

        let decimal_rule = ValidationRule::decimal(NumericConstraint::less_than(100.0));
        sheet.validations.set(CellRange::new(0, 2, 9, 2), decimal_rule);

        // Export
        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("roundtrip_multiple.xlsx");
        let export_result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(export_result.validations_exported, 3);

        // Import
        let (imported_workbook, import_result) = import(&export_path).unwrap();
        assert_eq!(import_result.validations_imported, 3);

        // Verify all rules exist
        let imported_sheet = imported_workbook.active_sheet();
        assert!(imported_sheet.validations.get(5, 0).is_some(), "List validation should exist");
        assert!(imported_sheet.validations.get(5, 1).is_some(), "Whole number validation should exist");
        assert!(imported_sheet.validations.get(5, 2).is_some(), "Decimal validation should exist");

        // Verify types
        assert!(matches!(
            imported_sheet.validations.get(5, 0).unwrap().rule_type,
            ValidationType::List(_)
        ));
        assert!(matches!(
            imported_sheet.validations.get(5, 1).unwrap().rule_type,
            ValidationType::WholeNumber(_)
        ));
        assert!(matches!(
            imported_sheet.validations.get(5, 2).unwrap().rule_type,
            ValidationType::Decimal(_)
        ));
    }

    // ========================================================================
    // Cell address conversion
    // ========================================================================

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(51), "AZ");
        assert_eq!(col_to_letter(52), "BA");
        assert_eq!(col_to_letter(701), "ZZ");
        assert_eq!(col_to_letter(702), "AAA");
    }

    #[test]
    fn test_cell_address() {
        assert_eq!(cell_address(0, 0), "A1");
        assert_eq!(cell_address(0, 1), "B1");
        assert_eq!(cell_address(4, 1), "B5");
        assert_eq!(cell_address(99, 25), "Z100");
        assert_eq!(cell_address(0, 26), "AA1");
    }

    // ========================================================================
    // Trust-critical golden tests: Date serial correctness
    // ========================================================================

    #[test]
    fn test_export_date_serial_known_values() {
        // These are known Excel serial values - if they're wrong, dates will be off
        use visigrid_engine::cell::{CellFormat, NumberFormat, DateStyle, date_to_serial};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Test critical date values (all verified against Excel)
        let test_dates = [
            // (row, serial, description)
            (0, 1.0, "Jan 1, 1900 = serial 1"),
            (1, 59.0, "Feb 28, 1900 = serial 59"),
            (2, 60.0, "Feb 29, 1900 (Excel's fake leap day) = serial 60"),
            (3, 61.0, "Mar 1, 1900 = serial 61"),
            (4, 36526.0, "Jan 1, 2000 = serial 36526"),
            (5, 45292.0, "Jan 1, 2024 = serial 45292"),
            (6, 45351.0, "Feb 29, 2024 (real leap day) = serial 45351"),
        ];

        for (row, serial, _desc) in test_dates.iter() {
            sheet.set_value(*row, 0, &serial.to_string());
            let mut format = CellFormat::default();
            format.number_format = NumberFormat::Date { style: DateStyle::Iso };
            sheet.set_format(*row, 0, format);
        }

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_dates.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 7);
        assert!(export_path.exists());

        // Verify roundtrip: the date functions should produce correct serials
        assert_eq!(date_to_serial(1900, 1, 1), 1.0);
        assert_eq!(date_to_serial(1900, 2, 28), 59.0);
        assert_eq!(date_to_serial(1900, 2, 29), 60.0); // Excel's fake day
        assert_eq!(date_to_serial(1900, 3, 1), 61.0);
        assert_eq!(date_to_serial(2000, 1, 1), 36526.0);
        assert_eq!(date_to_serial(2024, 1, 1), 45292.0);
        assert_eq!(date_to_serial(2024, 2, 29), 45351.0);
    }

    #[test]
    fn test_export_datetime_with_time_fraction() {
        use visigrid_engine::cell::{CellFormat, NumberFormat};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // DateTime = date serial + time fraction
        // Jan 1, 2024 at noon = 45292.5
        // Jan 1, 2024 at 6:00 AM = 45292.25
        // Jan 1, 2024 at 6:00 PM = 45292.75
        let test_datetimes = [
            (0, 45292.0, "Midnight"),
            (1, 45292.25, "6:00 AM"),
            (2, 45292.5, "Noon"),
            (3, 45292.75, "6:00 PM"),
            (4, 45292.999988425926, "23:59:59 (almost midnight)"),
        ];

        for (row, serial, _desc) in test_datetimes.iter() {
            sheet.set_value(*row, 0, &serial.to_string());
            let mut format = CellFormat::default();
            format.number_format = NumberFormat::DateTime;
            sheet.set_format(*row, 0, format);
        }

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_datetime.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 5);
    }

    // ========================================================================
    // Trust-critical golden tests: Number format fidelity
    // ========================================================================

    #[test]
    fn test_export_currency_formats() {
        use visigrid_engine::cell::{CellFormat, NumberFormat};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Currency with 0 decimals
        sheet.set_value(0, 0, "1234");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Currency { decimals: 0 };
        sheet.set_format(0, 0, format);

        // Currency with 2 decimals
        sheet.set_value(1, 0, "1234.56");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Currency { decimals: 2 };
        sheet.set_format(1, 0, format);

        // Negative currency
        sheet.set_value(2, 0, "-1234.56");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Currency { decimals: 2 };
        sheet.set_format(2, 0, format);

        // Zero currency
        sheet.set_value(3, 0, "0");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Currency { decimals: 2 };
        sheet.set_format(3, 0, format);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_currency.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 4);
    }

    #[test]
    fn test_export_percent_formats() {
        use visigrid_engine::cell::{CellFormat, NumberFormat};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Percent with 0 decimals (0.5 = 50%)
        sheet.set_value(0, 0, "0.5");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Percent { decimals: 0 };
        sheet.set_format(0, 0, format);

        // Percent with 2 decimals (0.1234 = 12.34%)
        sheet.set_value(1, 0, "0.1234");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Percent { decimals: 2 };
        sheet.set_format(1, 0, format);

        // 100% (1.0)
        sheet.set_value(2, 0, "1");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Percent { decimals: 0 };
        sheet.set_format(2, 0, format);

        // Over 100% (1.5 = 150%)
        sheet.set_value(3, 0, "1.5");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Percent { decimals: 0 };
        sheet.set_format(3, 0, format);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_percent.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 4);
    }

    #[test]
    fn test_export_large_numbers() {
        // Excel has 15-digit precision. Numbers larger than that lose precision.
        // Numbers >= 1e15 are exported as text to preserve exact value.
        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Safe: 15 digits (Excel can represent exactly) - exported as number
        sheet.set_value(0, 0, "123456789012345");

        // Risky: 16 digits - should be exported as text with warning
        sheet.set_value(1, 0, "1234567890123456");

        // At the boundary: 1e15 - should be exported as text
        sheet.set_value(2, 0, "1000000000000000");

        // Just under boundary: 999999999999999 - should be safe as number
        sheet.set_value(3, 0, "999999999999999");

        // Small precision number - safe
        sheet.set_value(4, 0, "0.000000001");

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_large_numbers.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 5);

        // Should have 2 precision warnings: rows 1 and 2 (>= 1e15)
        assert_eq!(result.precision_warnings.len(), 2);
        assert!(result.has_warnings());

        // Verify the addresses
        let addresses: Vec<_> = result.precision_warnings.iter().map(|w| w.address.as_str()).collect();
        assert!(addresses.contains(&"A2")); // 1234567890123456
        assert!(addresses.contains(&"A3")); // 1000000000000000
    }

    #[test]
    fn test_exceeds_excel_precision() {
        // Under the limit - safe
        assert!(!exceeds_excel_precision(999_999_999_999_999.0));
        assert!(!exceeds_excel_precision(123_456_789_012_345.0));
        assert!(!exceeds_excel_precision(0.0));
        assert!(!exceeds_excel_precision(-999_999_999_999_999.0));

        // At or over the limit - unsafe
        assert!(exceeds_excel_precision(1_000_000_000_000_000.0));
        assert!(exceeds_excel_precision(1_234_567_890_123_456.0));
        assert!(exceeds_excel_precision(-1_000_000_000_000_000.0));

        // Edge cases
        assert!(!exceeds_excel_precision(f64::NAN));
        assert!(!exceeds_excel_precision(f64::INFINITY));
    }

    #[test]
    fn test_export_number_edge_cases() {
        use visigrid_engine::cell::{CellFormat, NumberFormat};

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Negative zero
        sheet.set_value(0, 0, "-0");

        // Very small positive
        sheet.set_value(1, 0, "0.0000001");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Number { decimals: 7 };
        sheet.set_format(1, 0, format);

        // Negative with many decimals
        sheet.set_value(2, 0, "-123.456789");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::Number { decimals: 6 };
        sheet.set_format(2, 0, format);

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_edge_numbers.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.cells_exported, 3);
    }

    // ========================================================================
    // QA Fixtures for Manual Testing
    // Run with: cargo test -p visigrid-io -- qa_fixture --nocapture --ignored
    // Exports to: target/qa-fixtures/
    // ========================================================================

    fn qa_fixtures_dir() -> std::path::PathBuf {
        let dir = std::path::PathBuf::from("target/qa-fixtures");
        std::fs::create_dir_all(&dir).ok();
        dir
    }

    /// Invoice workbook: "Would I send this to a client?"
    /// Tests: currency, dates, bold headers, wrapped text, formulas, frozen panes
    #[test]
    #[ignore]
    fn qa_fixture_invoice() {
        use visigrid_engine::cell::{CellFormat, NumberFormat, DateStyle, Alignment, TextOverflow, date_to_serial};

        let mut workbook = Workbook::new();
        workbook.rename_sheet(0, "Invoice");
        let sheet = workbook.active_sheet_mut();

        // Row 0: Company header (bold, centered)
        sheet.set_value(0, 0, "ACME Corporation");
        let mut fmt = CellFormat::default();
        fmt.bold = true;
        fmt.alignment = Alignment::Center;
        sheet.set_format(0, 0, fmt);

        // Row 1: Invoice number
        sheet.set_value(1, 0, "Invoice #12345");

        // Row 3: Column headers (bold, wrapped)
        let headers = ["Item", "Description", "Qty", "Unit Price", "Total", "Date"];
        for (col, header) in headers.iter().enumerate() {
            sheet.set_value(3, col, header);
            let mut fmt = CellFormat::default();
            fmt.bold = true;
            fmt.text_overflow = TextOverflow::Wrap;
            sheet.set_format(3, col, fmt);
        }

        // Row 4-8: Line items
        let items = [
            ("Widget A", "Standard widget", 10, 25.00, "2024-01-15"),
            ("Widget B", "Premium widget with extended warranty", 5, 50.00, "2024-01-16"),
            ("Service Fee", "Installation", 1, 150.00, "2024-01-17"),
            ("Discount", "Volume discount", 1, -50.00, "2024-01-17"),
            ("Tax", "Sales tax 8%", 1, 0.0, "2024-01-17"),
        ];

        for (row_offset, (item, desc, qty, price, date)) in items.iter().enumerate() {
            let row = 4 + row_offset;
            sheet.set_value(row, 0, item);
            sheet.set_value(row, 1, desc);
            sheet.set_value(row, 2, &qty.to_string());
            sheet.set_value(row, 3, &price.to_string());

            // Total formula
            sheet.set_value(row, 4, &format!("=C{}*D{}", row + 1, row + 1));

            // Date (as serial)
            let parts: Vec<_> = date.split('-').collect();
            let serial = date_to_serial(
                parts[0].parse().unwrap(),
                parts[1].parse().unwrap(),
                parts[2].parse().unwrap(),
            );
            sheet.set_value(row, 5, &serial.to_string());

            // Format price column as currency
            let mut price_fmt = CellFormat::default();
            price_fmt.number_format = NumberFormat::Currency { decimals: 2 };
            sheet.set_format(row, 3, price_fmt.clone());
            sheet.set_format(row, 4, price_fmt);

            // Format date column
            let mut date_fmt = CellFormat::default();
            date_fmt.number_format = NumberFormat::Date { style: DateStyle::Iso };
            sheet.set_format(row, 5, date_fmt);
        }

        // Row 9: Totals
        sheet.set_value(9, 3, "Grand Total:");
        let mut fmt = CellFormat::default();
        fmt.bold = true;
        sheet.set_format(9, 3, fmt);

        sheet.set_value(9, 4, "=SUM(E5:E9)");
        let mut total_fmt = CellFormat::default();
        total_fmt.bold = true;
        total_fmt.number_format = NumberFormat::Currency { decimals: 2 };
        sheet.set_format(9, 4, total_fmt);

        // Export with layout
        let mut layout = ExportLayout::default();
        layout.col_widths.insert(0, 100.0);  // Item
        layout.col_widths.insert(1, 250.0);  // Description (wide)
        layout.col_widths.insert(2, 50.0);   // Qty
        layout.col_widths.insert(3, 100.0);  // Unit Price
        layout.col_widths.insert(4, 100.0);  // Total
        layout.col_widths.insert(5, 100.0);  // Date
        layout.frozen_rows = 4; // Freeze header rows

        let path = qa_fixtures_dir().join("invoice.xlsx");
        let result = export(&workbook, &path, Some(&[layout])).unwrap();

        println!("\n=== Invoice QA Fixture ===");
        println!("Exported to: {}", path.display());
        println!("Summary: {}", result.summary());
        println!("\nVerify in Excel:");
        println!("  [ ] Header row is bold");
        println!("  [ ] Currency shows $ with 2 decimals");
        println!("  [ ] Negative discount shows correctly");
        println!("  [ ] Dates show as YYYY-MM-DD");
        println!("  [ ] Formulas recalculate correctly");
        println!("  [ ] Row 4 frozen when scrolling");
        println!("  [ ] Column widths reasonable");
    }

    /// Operations report: "Does this feel like Excel?"
    /// Tests: percent, large numbers, alignment, multi-sheet
    #[test]
    #[ignore]
    fn qa_fixture_operations_report() {
        use visigrid_engine::cell::{CellFormat, NumberFormat, Alignment};

        let mut workbook = Workbook::new();
        workbook.rename_sheet(0, "Summary");

        // Summary sheet
        let summary = workbook.active_sheet_mut();
        summary.set_value(0, 0, "Operations Summary");
        let mut fmt = CellFormat::default();
        fmt.bold = true;
        summary.set_format(0, 0, fmt);

        summary.set_value(2, 0, "Metric");
        summary.set_value(2, 1, "Value");
        summary.set_value(2, 2, "Change");

        let mut header_fmt = CellFormat::default();
        header_fmt.bold = true;
        header_fmt.alignment = Alignment::Center;
        for col in 0..3 {
            summary.set_format(2, col, header_fmt.clone());
        }

        // Metrics with different formats
        let metrics = [
            ("Revenue", "1234567.89", "0.15"),     // Currency
            ("Expenses", "987654.32", "-0.08"),    // Currency, negative %
            ("Margin", "0.25", "0.02"),            // Percent
            ("Orders", "999999999999999", "0.05"), // Large but safe number
            ("ID (16-digit)", "1234567890123456", "0"),  // Should become text
        ];

        for (row_offset, (metric, value, change)) in metrics.iter().enumerate() {
            let row = 3 + row_offset;
            summary.set_value(row, 0, metric);
            summary.set_value(row, 1, value);
            summary.set_value(row, 2, change);

            // Percent format for change column
            let mut pct_fmt = CellFormat::default();
            pct_fmt.number_format = NumberFormat::Percent { decimals: 1 };
            pct_fmt.alignment = Alignment::Right;
            summary.set_format(row, 2, pct_fmt);
        }

        // Currency format for revenue/expenses
        let mut curr_fmt = CellFormat::default();
        curr_fmt.number_format = NumberFormat::Currency { decimals: 2 };
        summary.set_format(3, 1, curr_fmt.clone());
        summary.set_format(4, 1, curr_fmt);

        // Percent format for margin value
        let mut pct_fmt = CellFormat::default();
        pct_fmt.number_format = NumberFormat::Percent { decimals: 0 };
        summary.set_format(5, 1, pct_fmt);

        // Add Data sheet
        workbook.add_sheet();
        workbook.rename_sheet(1, "Data");
        let data = workbook.sheet_mut(1).unwrap();

        data.set_value(0, 0, "Raw Data");
        let mut fmt = CellFormat::default();
        fmt.bold = true;
        data.set_format(0, 0, fmt);

        data.set_value(1, 0, "Refer to Summary sheet for processed metrics");

        // Export
        let path = qa_fixtures_dir().join("operations_report.xlsx");
        let result = export(&workbook, &path, None).unwrap();

        println!("\n=== Operations Report QA Fixture ===");
        println!("Exported to: {}", path.display());
        println!("Summary: {}", result.summary());
        if result.has_warnings() {
            println!("Warnings: {}", result.warning_summary().unwrap_or_default());
            println!("\n{}", result.full_report());
        }
        println!("\nVerify in Excel:");
        println!("  [ ] Two sheets: Summary and Data");
        println!("  [ ] Percent changes show with % sign");
        println!("  [ ] 15-digit number displays correctly");
        println!("  [ ] 16-digit ID is text (check formula bar)");
        println!("  [ ] Negative percent shows correctly");
        println!("  [ ] Alignment is consistent");
    }

    /// Formula fallback workbook: "Prove honesty"
    /// Tests: invalid formulas, computed value fallback, detailed warnings
    #[test]
    #[ignore]
    fn qa_fixture_formula_fallback() {
        use visigrid_engine::cell::CellFormat;

        let mut workbook = Workbook::new();
        workbook.rename_sheet(0, "Formulas");
        let sheet = workbook.active_sheet_mut();

        // Header
        sheet.set_value(0, 0, "Formula Fallback Test");
        let mut fmt = CellFormat::default();
        fmt.bold = true;
        sheet.set_format(0, 0, fmt);

        sheet.set_value(2, 0, "Description");
        sheet.set_value(2, 1, "Formula/Value");
        sheet.set_value(2, 2, "Expected Result");
        for col in 0..3 {
            let mut fmt = CellFormat::default();
            fmt.bold = true;
            sheet.set_format(2, col, fmt);
        }

        // Test cases
        // Row 3: Valid formula (should export as formula)
        sheet.set_value(3, 0, "Valid SUM");
        sheet.set_value(3, 1, "=1+2+3");
        sheet.set_value(3, 2, "Should show =1+2+3 in formula bar, display 6");

        // Row 4: Another valid formula
        sheet.set_value(4, 0, "Valid IF");
        sheet.set_value(4, 1, "=IF(1>0,\"Yes\",\"No\")");
        sheet.set_value(4, 2, "Should show formula, display 'Yes'");

        // Row 5: Plain value
        sheet.set_value(5, 0, "Plain text");
        sheet.set_value(5, 1, "Hello World");
        sheet.set_value(5, 2, "Should be text");

        // Row 6: Number
        sheet.set_value(6, 0, "Plain number");
        sheet.set_value(6, 1, "42");
        sheet.set_value(6, 2, "Should be number 42");

        // Export
        let mut layout = ExportLayout::default();
        layout.col_widths.insert(0, 150.0);
        layout.col_widths.insert(1, 200.0);
        layout.col_widths.insert(2, 300.0);

        let path = qa_fixtures_dir().join("formula_fallback.xlsx");
        let result = export(&workbook, &path, Some(&[layout])).unwrap();

        println!("\n=== Formula Fallback QA Fixture ===");
        println!("Exported to: {}", path.display());
        println!("Summary: {}", result.summary());
        println!("Formulas exported: {}", result.formulas_exported);
        println!("Formulas as values: {}", result.formulas_as_values);
        if !result.converted_formulas.is_empty() {
            println!("\nConverted formulas:");
            println!("{}", result.converted_formulas_report());
        }
        println!("\nVerify in Excel:");
        println!("  [ ] Row 4 formula bar shows =1+2+3");
        println!("  [ ] Row 5 formula bar shows =IF(1>0,\"Yes\",\"No\")");
        println!("  [ ] Formulas recalculate correctly");
        println!("  [ ] Column widths are readable");
    }

    /// Generate all QA fixtures at once
    #[test]
    #[ignore]
    fn qa_fixture_generate_all() {
        println!("\n");
        println!("╔══════════════════════════════════════════════════════════════╗");
        println!("║                XLSX Export QA Fixtures                       ║");
        println!("╚══════════════════════════════════════════════════════════════╝");

        // Run each fixture generator (they're ignored, so call directly)
        // Note: This test just provides instructions since we can't call ignored tests
        println!("\nTo generate all fixtures, run:");
        println!("  cargo test -p visigrid-io -- qa_fixture --nocapture --ignored");
        println!("\nFixtures will be created in: target/qa-fixtures/");
        println!("\nFiles to test:");
        println!("  1. invoice.xlsx - Currency, dates, formulas, frozen panes");
        println!("  2. operations_report.xlsx - Percent, large numbers, multi-sheet");
        println!("  3. formula_fallback.xlsx - Formula export verification");
        println!("\nTest each file in:");
        println!("  [ ] Excel on Windows");
        println!("  [ ] Excel on Mac");
        println!("  [ ] LibreOffice Calc");
        println!("  [ ] Google Sheets (File > Import)");
    }

    // ========================================================================
    // Import performance benchmarks
    // Run with: cargo test -p visigrid-io --release -- import_benchmark --nocapture --ignored
    // ========================================================================

    fn run_import_benchmark(path: &str) {
        let path = std::path::Path::new(path);
        if !path.exists() {
            println!("  Skipped: {} (file not found)", path.display());
            println!("  Generate fixtures with: python3 benchmarks/generate_xlsx_fixtures.py");
            return;
        }

        let file_size = std::fs::metadata(path).map(|m| m.len()).unwrap_or(0);

        match import(path) {
            Ok((_workbook, result)) => {
                println!("  File: {} ({:.1} KB)", path.file_name().unwrap().to_string_lossy(), file_size as f64 / 1024.0);
                println!("  Cells: {:>10}", result.cells_imported);
                println!("  Formulas: {:>7}", result.formulas_imported);
                println!("  Duration: {:>6} ms", result.import_duration_ms);
                println!("  Rate: {:>10.0} cells/sec", result.cells_imported as f64 / (result.import_duration_ms as f64 / 1000.0));
                println!();
            }
            Err(e) => {
                println!("  Error: {}", e);
            }
        }
    }

    #[test]
    #[ignore]  // Run explicitly with --ignored flag
    fn import_benchmark_small() {
        println!();
        println!("=== Small file benchmark ===");
        run_import_benchmark("../../benchmarks/fixtures/small.xlsx");
    }

    #[test]
    #[ignore]
    fn import_benchmark_medium() {
        println!();
        println!("=== Medium file benchmark ===");
        run_import_benchmark("../../benchmarks/fixtures/medium.xlsx");
    }

    #[test]
    #[ignore]
    fn import_benchmark_large() {
        println!();
        println!("=== Large file benchmark ===");
        run_import_benchmark("../../benchmarks/fixtures/large.xlsx");
    }

    #[test]
    #[ignore]
    fn import_benchmark_all() {
        println!();
        println!("╔══════════════════════════════════════════════════════════════╗");
        println!("║              XLSX Import Performance Benchmark               ║");
        println!("╚══════════════════════════════════════════════════════════════╝");
        println!();

        println!("Small (~5k cells):");
        run_import_benchmark("../../benchmarks/fixtures/small.xlsx");

        println!("Medium (~200k cells with formulas):");
        run_import_benchmark("../../benchmarks/fixtures/medium.xlsx");

        println!("Large (~1M cells):");
        run_import_benchmark("../../benchmarks/fixtures/large.xlsx");

        println!("Decision guide:");
        println!("  - Small <50ms, Medium <250ms: Background import not needed");
        println!("  - Medium >250ms or Large freezes: Implement background import");
    }
}
