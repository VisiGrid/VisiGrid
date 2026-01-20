// Excel file import (xlsx, xls, xlsb, ods)
//
// This is a one-way import. Files are converted to VisiGrid's internal model.
// Features that don't map cleanly are dropped with warnings.

use std::collections::HashMap;
use std::path::Path;
use std::time::Instant;

use calamine::{open_workbook_auto, Data, Reader, Sheets};
use visigrid_engine::cell::{DateStyle, NumberFormat};
use visigrid_engine::formula::analyze::tally_unknown_functions;
use visigrid_engine::formula::parser::parse as parse_formula;
use visigrid_engine::sheet::Sheet;
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
            let mut sheet = Sheet::new(MAX_ROWS, MAX_COLS);
            sheet.name = sheet_name.clone();
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

        let mut sheet = Sheet::new(MAX_ROWS, MAX_COLS);
        sheet.name = sheet_name.clone();

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

        sheets.push(sheet);
        result.sheets_imported += 1;
        result.sheet_stats.push(stats);
    }

    if sheets.is_empty() {
        return Err("No sheets could be imported".to_string());
    }

    // Record import duration
    result.import_duration_ms = start_time.elapsed().as_millis();

    let workbook = Workbook::from_sheets(sheets, 0);
    Ok((workbook, result))
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
