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
use visigrid_engine::formula::eval::Value;
use visigrid_engine::formula::parser::parse as parse_formula;
use visigrid_engine::sheet::{MergedRegion, Sheet, SheetId};
use visigrid_engine::workbook::Workbook;
use crate::xlsx_styles;

/// A concrete example of a post-recalc error, for the import report
#[derive(Debug, Clone)]
pub struct RecalcErrorExample {
    /// Sheet name where the error occurred
    pub sheet: String,
    /// Cell address (e.g., "B5")
    pub address: String,
    /// Error category: "circular" or "error"
    pub kind: &'static str,
    /// The error message (e.g., "#REF!", "#CYCLE!")
    pub error: String,
    /// The original formula, if available
    pub formula: Option<String>,
}

/// Maximum number of error examples to collect for the import report
const MAX_ERROR_EXAMPLES: usize = 5;

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
    pub recalc_errors: usize,             // Formula errors after recalc
    pub recalc_circular: usize,           // Circular reference errors after recalc
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
    /// Total formula errors after recalc (Value::Error from evaluation)
    pub recalc_errors: usize,
    /// Total circular reference errors after recalc
    pub recalc_circular: usize,
    /// Number of shared formula groups detected in XLSX XML
    pub shared_formula_groups: usize,
    /// Sample formulas that failed to parse (for diagnostics, up to 10)
    pub parse_error_samples: Vec<String>,
    /// Top error examples for the import report (up to MAX_ERROR_EXAMPLES)
    pub recalc_error_examples: Vec<RecalcErrorExample>,
    /// Number of cells that received a style_id from XLSX formatting
    pub styles_imported: usize,
    /// Number of unique styles in the workbook style table
    pub unique_styles: usize,
    /// Unsupported formatting features encountered during import
    pub unsupported_format_features: Vec<String>,
    /// Per-sheet imported layout (column widths, row heights) in raw Excel units.
    /// Indexed by sheet position (same order as sheets in workbook).
    pub imported_layouts: Vec<ImportedLayout>,
    /// Formula cells in XLSX XML that had no cached <v> value (calamine may skip these)
    pub formula_cells_without_values: usize,
    /// Value cells backfilled from XLSX XML (shared strings, inline strings, numbers calamine missed)
    pub value_cells_backfilled: usize,
    /// Total merged cell regions imported
    pub merges_imported: usize,
    /// Merged regions dropped due to overlap with existing merges
    pub merges_dropped_overlap: usize,
    /// Merged regions dropped due to invalid cell references
    pub merges_dropped_invalid: usize,
    /// Cycle cells frozen to cached values (freeze_cycles option)
    pub cycles_frozen: usize,
    /// Cycle cells with no cached value (remain #CYCLE!)
    pub cycles_no_cached: usize,
    /// Whether freeze was applied (freeze_cycles requested AND cells were frozen)
    pub freeze_applied: bool,
    /// Formula strings collected during values_only import.
    /// Key: (sheet_index, row, col), Value: formula string (with leading =).
    /// Empty unless values_only is true.
    pub formula_strings: HashMap<(usize, usize, usize), String>,
}

/// Column/row dimension data imported from XLSX, in raw Excel units.
#[derive(Debug, Default, Clone)]
pub struct ImportedLayout {
    /// Column index → raw Excel character-width units
    pub col_widths: HashMap<usize, f64>,
    /// Row index → raw Excel point units
    pub row_heights: HashMap<usize, f64>,
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
        if self.styles_imported > 0 {
            parts.push(format!("Formatting: {} cells ({} styles)",
                self.styles_imported, self.unique_styles));
        }
        if self.merges_imported > 0 {
            let dropped = self.merges_dropped_overlap + self.merges_dropped_invalid;
            if dropped > 0 {
                parts.push(format!("{} merged regions ({} dropped)", self.merges_imported, dropped));
            } else {
                parts.push(format!("{} merged regions", self.merges_imported));
            }
        }
        parts.join(" · ")
    }

    /// Returns true if there are actionable warnings
    pub fn has_warnings(&self) -> bool {
        self.truncated
            || self.formulas_failed > 0
            || !self.unsupported_functions.is_empty()
            || !self.warnings.is_empty()
            || self.recalc_errors > 0
            || self.recalc_circular > 0
            || self.merges_dropped_overlap > 0
            || self.merges_dropped_invalid > 0
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

        if self.recalc_errors > 0 {
            issues.push(format!("{} formula errors", self.recalc_errors));
        }
        if self.recalc_circular > 0 {
            issues.push(format!("{} circular references", self.recalc_circular));
        }

        let merges_dropped = self.merges_dropped_overlap + self.merges_dropped_invalid;
        if merges_dropped > 0 {
            issues.push(format!("{} merged regions dropped", merges_dropped));
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

/// Import options controlling optional behavior during Excel import.
#[derive(Clone)]
pub struct ImportOptions {
    /// If true, cycle cells are frozen to their cached values from the XLSX file.
    /// This makes workbooks with circular references (e.g. iterative calculation
    /// models from Excel) immediately usable.
    pub freeze_cycles: bool,

    /// Enable iterative calculation during import recalc.
    /// When true, SCCs are resolved via Jacobi iteration instead of marking #CYCLE!.
    pub iterative_enabled: bool,

    /// Maximum iterations per SCC (used when iterative_enabled is true).
    pub iterative_max_iters: u32,

    /// Convergence tolerance (used when iterative_enabled is true).
    pub iterative_tolerance: f64,

    /// When true, preserve cached cell values instead of replacing them with formulas.
    /// Formula strings are collected in ImportResult.formula_strings instead.
    /// Skips dependency graph rebuild and formula recomputation.
    /// Cell values remain as calamine extracted them (strings, numbers, dates as serials).
    pub values_only: bool,
}

impl Default for ImportOptions {
    fn default() -> Self {
        Self {
            freeze_cycles: false,
            iterative_enabled: false,
            iterative_max_iters: 100,
            iterative_tolerance: 1e-9,
            values_only: false,
        }
    }
}

/// Maximum number of cells to import (prevents DoS from huge files)
const MAX_CELLS: usize = 5_000_000;

/// Maximum dimensions for a sheet
const MAX_ROWS: usize = 65536;
const MAX_COLS: usize = 256;

/// Import an Excel file (xlsx, xls, xlsb, ods)
pub fn import(path: &Path) -> Result<(Workbook, ImportResult), String> {
    import_with_options(path, &ImportOptions::default())
}

/// Import an Excel file with options (xlsx, xls, xlsb, ods)
pub fn import_with_options(path: &Path, options: &ImportOptions) -> Result<(Workbook, ImportResult), String> {
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

    // Per-sheet snapshot of typed cached values for cycle freeze.
    // Key: (row, col), Value: (cached_value_or_none, formula_source)
    // None means "Calamine had no cached result at all" (distinct from
    // Some(CellValue::Empty) which means "Calamine had an explicit empty value").
    let mut cached_snapshots: Vec<HashMap<(usize, usize), (Option<CellValue>, String)>> = Vec::new();

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

            cached_snapshots.push(HashMap::new());
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
        cached_snapshots.push(HashMap::new());

        // Range start offset (data may not begin at A1)
        let (data_start_row, data_start_col) = range.start().unwrap_or((0, 0));

        for (row_idx, row) in range.rows().enumerate() {
            let target_row = data_start_row as usize + row_idx;
            if target_row >= effective_rows {
                break;
            }

            for (col_idx, cell) in row.iter().enumerate() {
                let target_col = data_start_col as usize + col_idx;
                if target_col >= effective_cols {
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

                match cell {
                    Data::Empty => {
                        // Skip empty cells. For freeze_cycles, the pass-2 snapshot
                        // will see None (no cell) for these positions. This is
                        // conservative: a formula whose cached result was truly
                        // empty gets counted as cycles_no_cached rather than
                        // frozen-as-empty. Acceptable trade-off vs materializing
                        // millions of phantom cells in the dense calamine range.
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
                // Formula range may start at a different offset than data range
                let (formula_start_row, formula_start_col) = formula_range.start().unwrap_or((0, 0));

                for (row_idx, row) in formula_range.rows().enumerate() {
                    let target_row = formula_start_row as usize + row_idx;
                    if target_row >= effective_rows {
                        break;
                    }

                    for (col_idx, formula) in row.iter().enumerate() {
                        let target_col = formula_start_col as usize + col_idx;
                        if target_col >= effective_cols {
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
                            // Strip ODS OpenFormula namespace prefix (e.g. "=of:SUM()" → "=SUM()")
                            let formula_str = strip_ods_prefix(&formula_str);

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
                                    if result.parse_error_samples.len() < 10 {
                                        result.parse_error_samples.push(format!(
                                            "{}!{}: {}",
                                            sheet_name,
                                            cell_address(target_row, target_col),
                                            formula_str
                                        ));
                                    }
                                }
                            }

                            if options.values_only {
                                // Preserve cached value; stash formula string for inspect
                                let sheet_idx = sheets.len();
                                result.formula_strings.insert(
                                    (sheet_idx, target_row, target_col),
                                    formula_str.clone(),
                                );
                            } else {
                                // Snapshot the typed cached value before formula overwrite
                                if options.freeze_cycles {
                                    let cached = sheet.get_cell_opt(target_row, target_col)
                                        .map(|c| c.value.clone());
                                    cached_snapshots.last_mut().unwrap().insert(
                                        (target_row, target_col),
                                        (cached, formula_str.clone()),
                                    );
                                }

                                sheet.set_value(target_row, target_col, &formula_str);
                                sheet.set_format(target_row, target_col, existing_format);
                            }

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

    // Import formatting from styles.xml and per-cell style IDs
    import_formatting(path, &sheet_names, &mut workbook, &mut result);

    if !options.values_only {
        // Detect shared formula groups from XLSX XML (diagnostic guardrail)
        result.shared_formula_groups = count_shared_formula_groups(path);

        // Extract formula-only cells from worksheet XML (cells with <f> but possibly no <v>).
        // calamine may skip these, causing "phantom blank" cells where formulas should exist.
        let xml_formulas = extract_xml_formula_cells(path);
        let mut formula_backfill_count = 0usize;
        for (sheet_idx, row, col, formula_text) in &xml_formulas {
            if let Some(sheet) = workbook.sheet_mut(*sheet_idx) {
                // Only backfill if the cell is currently empty (calamine didn't import it)
                if sheet.get_raw(*row, *col).is_empty() {
                    let formula_str = if formula_text.starts_with('=') {
                        formula_text.clone()
                    } else {
                        format!("={}", formula_text)
                    };
                    eprintln!("[XLSX backfill] {}{}: ={}",
                        col_to_letter(*col), *row + 1, formula_text);
                    sheet.set_value(*row, *col, &formula_str);
                    formula_backfill_count += 1;
                }
            }
        }
        result.formula_cells_without_values = formula_backfill_count;
        if formula_backfill_count > 0 {
            eprintln!("[XLSX import] Backfilled {} formula-only cells from XML ({} total XML formulas parsed)",
                formula_backfill_count, xml_formulas.len());
        }

        // Backfill value-only cells from XLSX XML.
        // calamine may skip cells stored as shared strings, inline strings, or numbers
        // that it doesn't surface through its cell iterator.
        let xml_values = extract_xml_value_cells(path);
        let mut value_backfill_count = 0usize;
        for (sheet_idx, row, col, value_text) in &xml_values {
            if let Some(sheet) = workbook.sheet_mut(*sheet_idx) {
                if sheet.get_raw(*row, *col).is_empty() {
                    eprintln!("[XLSX backfill] {}{}: value=\"{}\"",
                        col_to_letter(*col), *row + 1, value_text);
                    sheet.set_value(*row, *col, value_text);
                    value_backfill_count += 1;
                }
            }
        }
        result.value_cells_backfilled = value_backfill_count;
        if value_backfill_count > 0 {
            eprintln!("[XLSX import] Backfilled {} value cells from XML ({} total XML value cells parsed)",
                value_backfill_count, xml_values.len());
        }

        // Rebuild dependency graph after loading all data
        workbook.rebuild_dep_graph();

        // Freeze cycle cells if requested
        if options.freeze_cycles {
            let cycle_members = workbook.dep_graph().find_cycle_members();
            if !cycle_members.is_empty() {
                // Build sheet ID → index mapping
                let sheet_id_to_idx: HashMap<SheetId, usize> = workbook.sheets()
                    .iter()
                    .enumerate()
                    .map(|(idx, s)| (s.id, idx))
                    .collect();

                let mut frozen_count = 0usize;
                let mut cycles_no_cached = 0usize;

                for cell_id in &cycle_members {
                    let sheet_idx = match sheet_id_to_idx.get(&cell_id.sheet) {
                        Some(&idx) => idx,
                        None => continue,
                    };

                    match cached_snapshots.get_mut(sheet_idx)
                        .and_then(|snap| snap.remove(&(cell_id.row, cell_id.col)))
                    {
                        Some((Some(cached_cv), formula_src)) => {
                            workbook.sheets_mut()[sheet_idx]
                                .freeze_cell(cell_id.row, cell_id.col, cached_cv, formula_src);
                            frozen_count += 1;
                        }
                        Some((None, _)) | None => {
                            // No cached result — cell will get #CYCLE! during recalc
                            cycles_no_cached += 1;
                        }
                    }
                }

                if frozen_count > 0 {
                    // Rebuild dep graph — frozen cells are no longer formula cells
                    workbook.rebuild_dep_graph();
                }

                result.cycles_frozen = frozen_count;
                result.cycles_no_cached = cycles_no_cached;
                result.freeze_applied = frozen_count > 0;

                eprintln!("[XLSX import] Froze {} cycle cells ({} without cached values)",
                    frozen_count, cycles_no_cached);
            }
        }

        // Wire iteration settings before recalc (if requested)
        if options.iterative_enabled {
            workbook.set_iterative_enabled(true);
            workbook.set_iterative_max_iters(options.iterative_max_iters);
            workbook.set_iterative_tolerance(options.iterative_tolerance);
        }

        // Recompute all formulas in topological order.
        // Individual set_value() calls during import evaluated formulas at sheet level
        // without proper dependency ordering — upstream cells may not have existed yet.
        // This full ordered recompute clears stale caches and evaluates everything correctly.
        let recalc_report = workbook.recompute_full_ordered();
        eprintln!("[XLSX import] Recomputed {} formulas in topo order (cycles: {})",
            recalc_report.cells_recomputed, recalc_report.had_cycles);

        // Post-recalc error counting: detect circular refs and formula evaluation errors
        for (sheet_idx, sheet) in workbook.sheets().iter().enumerate() {
            let mut sheet_errors = 0usize;
            let mut sheet_circular = 0usize;
            for ((_row, _col), cell) in sheet.cells_iter() {
                // Circulars: structural graph property (set during dep graph cycle detection)
                if cell.value.is_cycle_error() {
                    sheet_circular += 1;
                    if result.recalc_error_examples.len() < MAX_ERROR_EXAMPLES {
                        result.recalc_error_examples.push(RecalcErrorExample {
                            sheet: sheet.name.clone(),
                            address: cell_address(*_row, *_col),
                            kind: "circular",
                            error: "#CYCLE!".to_string(),
                            formula: None, // Source is lost when cycle is detected
                        });
                    }
                    continue;
                }
                // Formula errors: evaluate formula cells, check for Value::Error
                if cell.value.formula_ast().is_some() {
                    if let Value::Error(ref e) = sheet.get_computed_value(*_row, *_col) {
                        sheet_errors += 1;
                        if result.recalc_error_examples.len() < MAX_ERROR_EXAMPLES {
                            let formula_source = match &cell.value {
                                CellValue::Formula { source, .. } => Some(source.clone()),
                                _ => None,
                            };
                            result.recalc_error_examples.push(RecalcErrorExample {
                                sheet: sheet.name.clone(),
                                address: cell_address(*_row, *_col),
                                kind: "error",
                                error: e.clone(),
                                formula: formula_source,
                            });
                        }
                    }
                }
            }
            if sheet_idx < result.sheet_stats.len() {
                result.sheet_stats[sheet_idx].recalc_errors = sheet_errors;
                result.sheet_stats[sheet_idx].recalc_circular = sheet_circular;
            }
            result.recalc_errors += sheet_errors;
            result.recalc_circular += sheet_circular;
        }
    }

    Ok((workbook, result))
}

// =============================================================================
// XLSX Formatting Import
// =============================================================================

/// Import formatting (styles, column widths, row heights) from XLSX into the workbook.
/// This is called after calamine has imported cell data, to layer formatting on top.
fn import_formatting(
    path: &Path,
    sheet_names: &[String],
    workbook: &mut Workbook,
    result: &mut ImportResult,
) {
    // Parse styles.xml and per-sheet formatting from the XLSX ZIP
    let (style_table, sheet_formats, stats) = match xlsx_styles::parse_xlsx_formatting(path, sheet_names) {
        Ok(data) => data,
        Err(_) => return, // Graceful fallback: no formatting
    };

    if style_table.len() == 0 {
        return; // No styles to apply
    }

    // Build a mapping from xlsx style index → workbook style_table index
    // by interning each parsed style into the workbook's global table
    let mut style_id_map: Vec<Option<u32>> = Vec::with_capacity(style_table.len());
    for style in &style_table.styles {
        if *style == CellFormat::default() {
            style_id_map.push(None); // Default style, no need to store
        } else {
            style_id_map.push(Some(workbook.intern_style(style.clone())));
        }
    }

    // Apply per-cell style IDs and handle styled-empty cells
    for (sheet_idx, sheet_fmt) in sheet_formats.iter().enumerate() {
        let sheet = match workbook.sheet_mut(sheet_idx) {
            Some(s) => s,
            None => continue,
        };

        for &(row, col, xlsx_style_id) in &sheet_fmt.cell_styles {
            // Look up the workbook style_id for this xlsx style index
            let wb_style_id = match style_id_map.get(xlsx_style_id) {
                Some(Some(id)) => *id,
                _ => continue, // Default style or out of range
            };

            // Get the resolved format for this style
            let resolved_format = match style_table.get(xlsx_style_id) {
                Some(f) => f,
                None => continue,
            };

            // Check if cell already has data (from calamine import)
            let cell_exists = !matches!(sheet.get_cell(row, col).value, CellValue::Empty);

            if cell_exists {
                // Cell has data: apply the style and set format
                sheet.set_style_id(row, col, wb_style_id);
                sheet.set_format_from_import(row, col, resolved_format.clone());
                result.styles_imported += 1;
            } else if xlsx_styles::is_style_visually_relevant(resolved_format) {
                // Styled-empty cell with visual formatting: materialize it
                sheet.set_style_id(row, col, wb_style_id);
                sheet.set_format_from_import(row, col, resolved_format.clone());
                result.styles_imported += 1;
            }
            // Else: styled-empty with no visual formatting → skip
        }
    }

    // Import merged cell regions
    for (sheet_idx, sheet_fmt) in sheet_formats.iter().enumerate() {
        if sheet_fmt.merged_regions.is_empty() {
            continue;
        }
        let sheet = match workbook.sheet_mut(sheet_idx) {
            Some(s) => s,
            None => continue,
        };
        for &(sr, sc, er, ec) in &sheet_fmt.merged_regions {
            // Validate ref bounds
            if sr > er || sc > ec {
                result.merges_dropped_invalid += 1;
                continue;
            }
            let region = MergedRegion::new(sr, sc, er, ec);
            match sheet.add_merge(region) {
                Ok(()) => result.merges_imported += 1,
                Err(msg) => {
                    result.merges_dropped_overlap += 1;
                    result.unsupported_format_features.push(format!("dropped merge: {}", msg));
                }
            }
        }
    }

    result.unique_styles = workbook.style_table.len();
    result.unsupported_format_features.extend(stats.unsupported_features);

    // Add unsupported features as warnings
    for feature in &result.unsupported_format_features {
        result.warnings.push(format!("Unsupported formatting: {}", feature));
    }

    // Collect layout data for app-level application
    result.imported_layouts = sheet_formats
        .into_iter()
        .map(|sf| ImportedLayout {
            col_widths: sf.col_widths,
            row_heights: sf.row_heights,
        })
        .collect();
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
    /// Total merged cell regions exported
    pub merges_exported: usize,
    /// Whether autofilter was exported
    pub autofilter_exported: bool,
    /// Number of hidden rows exported
    pub hidden_rows_exported: usize,
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
    /// AutoFilter range: (min_row, min_col, max_row, max_col)
    pub autofilter_range: Option<(usize, usize, usize, usize)>,
    /// Hidden data rows (from filter visibility mask)
    pub hidden_rows: Vec<usize>,
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

        // Export merged cell regions first — merge_range() writes blanks to all
        // cells in the range, then export_sheet_cells() overwrites the origin
        // cell with the correct typed value (number, formula, etc.).
        let merge_format = Format::new();
        for merge in &sheet.merged_regions {
            worksheet
                .merge_range(
                    merge.start.0 as u32,
                    merge.start.1 as u16,
                    merge.end.0 as u32,
                    merge.end.1 as u16,
                    "",
                    &merge_format,
                )
                .map_err(|e| format!("Failed to write merge: {}", e))?;
            result.merges_exported += 1;
        }

        // Export cells (skips merge-hidden cells; origin cells overwrite the
        // blank written by merge_range above)
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

        // Export autofilter state
        if let Some(layout) = layout {
            if let Some((min_r, min_c, max_r, max_c)) = layout.autofilter_range {
                worksheet
                    .autofilter(min_r as u32, min_c as u16, max_r as u32, max_c as u16)
                    .map_err(|e| format!("Failed to set autofilter: {}", e))?;
                result.autofilter_exported = true;
            }

            for &row in &layout.hidden_rows {
                worksheet
                    .set_row_hidden(row as u32)
                    .map_err(|e| format!("Failed to hide row {}: {}", row, e))?;
            }
            result.hidden_rows_exported += layout.hidden_rows.len();
        }

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

/// Strip ODS OpenFormula namespace prefixes and convert argument separators.
///
/// ODS files use prefixed function names like `=of:SUM(A1:A10)` or `=of:IF(...)`,
/// and use semicolons (`;`) as argument separators instead of commas.
/// This converts them to standard `=SUM(A1:A10)` syntax with comma separators.
fn strip_ods_prefix(formula: &str) -> String {
    if !formula.starts_with("=of:") && !formula.starts_with("=OF:") {
        return formula.to_string();
    }

    let mut result = String::with_capacity(formula.len());
    result.push('=');
    let body = &formula[1..]; // skip leading '='
    let bytes = body.as_bytes();
    let mut i = 0;
    let mut in_string = false;

    while i < bytes.len() {
        // Track string literals to avoid converting semicolons inside them
        if bytes[i] == b'"' {
            in_string = !in_string;
            result.push('"');
            i += 1;
            continue;
        }

        if in_string {
            result.push(bytes[i] as char);
            i += 1;
            continue;
        }

        // Strip "of:" prefix (case-insensitive)
        if i + 3 <= bytes.len()
            && (bytes[i] == b'o' || bytes[i] == b'O')
            && (bytes[i + 1] == b'f' || bytes[i + 1] == b'F')
            && bytes[i + 2] == b':'
        {
            i += 3;
        // Convert ODS semicolon argument separator to comma
        } else if bytes[i] == b';' {
            result.push(',');
            i += 1;
        } else {
            result.push(bytes[i] as char);
            i += 1;
        }
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

        // Skip merge-hidden cells - only the origin cell exports its value
        if sheet.is_merge_hidden(*row, *col) {
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

/// Count shared formula groups in an XLSX file by scanning worksheet XML.
///
/// Shared formulas use `<f t="shared" ref="..." si="N">` master nodes.
/// This function counts master nodes (those with both `t="shared"` and `ref=...`)
/// across all worksheets. This serves as a diagnostic guardrail: if calamine
/// doesn't expand shared formulas correctly, this count helps diagnose the issue.
///
/// Returns 0 for non-XLSX formats (xls, xlsb, ods) or on any error.
fn count_shared_formula_groups(path: &Path) -> usize {
    use zip::ZipArchive;

    // Only works for XLSX files (ZIP-based XML)
    let file = match std::fs::File::open(path) {
        Ok(f) => f,
        Err(_) => return 0,
    };
    let mut archive = match ZipArchive::new(file) {
        Ok(a) => a,
        Err(_) => return 0,
    };

    // Parse workbook.xml to get sheet rIds
    let workbook_xml = match read_zip_file_for_shared(&mut archive, "xl/workbook.xml") {
        Some(s) => s,
        None => return 0,
    };
    let rels_xml = match read_zip_file_for_shared(&mut archive, "xl/_rels/workbook.xml.rels") {
        Some(s) => s,
        None => return 0,
    };

    // Collect all worksheet XML paths
    let worksheet_paths = resolve_worksheet_paths(&workbook_xml, &rels_xml);

    let mut total_groups = 0;

    for ws_path in worksheet_paths {
        let xml = match read_zip_file_for_shared(&mut archive, &ws_path) {
            Some(s) => s,
            None => continue,
        };
        total_groups += count_shared_masters_in_xml(&xml);
    }

    total_groups
}

/// Read a file from a ZIP archive, returning None on error.
fn read_zip_file_for_shared<R: std::io::Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Option<String> {
    use std::io::Read;
    let mut file = archive.by_name(path).ok()?;
    let mut content = String::new();
    file.read_to_string(&mut content).ok()?;
    Some(content)
}

/// Resolve worksheet XML paths from workbook.xml + workbook.xml.rels
fn resolve_worksheet_paths(workbook_xml: &str, rels_xml: &str) -> Vec<String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut paths = Vec::new();

    // Step 1: Parse workbook.xml to get sheet rIds
    let mut rids = Vec::new();
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.name().as_ref() == b"sheet" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"r:id" {
                        rids.push(String::from_utf8_lossy(&attr.value).to_string());
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Step 2: Parse rels to resolve rId → target path
    let mut rid_to_target: HashMap<String, String> = HashMap::new();
    let mut reader = Reader::from_str(rels_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = Some(String::from_utf8_lossy(&attr.value).to_string()),
                        b"Target" => target = Some(String::from_utf8_lossy(&attr.value).to_string()),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    rid_to_target.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Step 3: Resolve each rId to a full path
    for rid in &rids {
        if let Some(target) = rid_to_target.get(rid) {
            // Only include worksheet targets (skip chartsheets, etc.)
            if target.contains("worksheet") {
                paths.push(format!("xl/{}", target));
            }
        }
    }

    paths
}

/// Count shared formula master nodes in a single worksheet XML.
/// Masters have both `t="shared"` and `ref="..."` attributes on `<f>` elements.
fn count_shared_masters_in_xml(xml: &str) -> usize {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut count = 0;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) if e.name().as_ref() == b"f" => {
                let mut is_shared = false;
                let mut has_ref = false;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"t" if attr.value.as_ref() == b"shared" => is_shared = true,
                        b"ref" => has_ref = true,
                        _ => {}
                    }
                }
                if is_shared && has_ref {
                    count += 1;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    count
}

/// Extract all formula cells from worksheet XML, including those with no cached <v> value.
/// Returns: Vec<(sheet_index, row, col, formula_text)>
/// sheet_index corresponds to the order in which sheets appear in workbook.xml.
fn extract_xml_formula_cells(path: &Path) -> Vec<(usize, usize, usize, String)> {
    use zip::ZipArchive;

    let file = match std::fs::File::open(path) {
        Ok(f) => f,
        Err(_) => return Vec::new(),
    };
    let mut archive = match ZipArchive::new(file) {
        Ok(a) => a,
        Err(_) => return Vec::new(),
    };

    let workbook_xml = match read_zip_file_for_shared(&mut archive, "xl/workbook.xml") {
        Some(s) => s,
        None => return Vec::new(),
    };
    let rels_xml = match read_zip_file_for_shared(&mut archive, "xl/_rels/workbook.xml.rels") {
        Some(s) => s,
        None => return Vec::new(),
    };

    let worksheet_paths = resolve_worksheet_paths(&workbook_xml, &rels_xml);

    let mut all_formulas = Vec::new();

    for (sheet_idx, ws_path) in worksheet_paths.iter().enumerate() {
        let xml = match read_zip_file_for_shared(&mut archive, ws_path) {
            Some(s) => s,
            None => continue,
        };
        extract_formulas_from_xml(&xml, sheet_idx, &mut all_formulas);
    }

    all_formulas
}

/// Shared formula definition collected from worksheet XML.
struct SharedFormulaDef {
    /// The base cell (row, col) where the master formula lives
    base_row: usize,
    base_col: usize,
    /// The master formula text (without '=')
    formula: String,
}

/// Parse a single worksheet XML and extract all formula cells, including
/// shared formula followers that have no formula text.
///
/// 2-pass approach:
/// Pass 1: Collect all <f> elements — masters have formula text, followers have si only.
///         Build shared formula definition map (si -> SharedFormulaDef).
/// Pass 2: For each follower, compute the row/col offset from the master's base cell,
///         and shift the formula references accordingly.
fn extract_formulas_from_xml(
    xml: &str,
    sheet_idx: usize,
    out: &mut Vec<(usize, usize, usize, String)>,
) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    // Intermediate struct for raw parsed formula entries
    struct FormulaEntry {
        row: usize,
        col: usize,
        formula_text: Option<String>,   // Some for masters, None for followers
        shared_si: Option<String>,       // shared index if t="shared"
    }

    let mut entries: Vec<FormulaEntry> = Vec::new();
    let mut shared_defs: HashMap<String, SharedFormulaDef> = HashMap::new();

    // --- Pass 1: Parse all <f> elements ---
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut current_cell_ref: Option<String> = None;
    // Track formula element attributes for the current <f>
    let mut in_formula = false;
    let mut current_f_shared: bool = false;
    let mut current_f_si: Option<String> = None;
    let mut current_f_has_ref: bool = false;
    let mut current_f_text: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"c" => {
                        current_cell_ref = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"r" {
                                current_cell_ref = Some(
                                    String::from_utf8_lossy(&attr.value).to_string()
                                );
                            }
                        }
                    }
                    b"f" => {
                        in_formula = true;
                        current_f_shared = false;
                        current_f_si = None;
                        current_f_has_ref = false;
                        current_f_text = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"t" if attr.value.as_ref() == b"shared" => {
                                    current_f_shared = true;
                                }
                                b"si" => {
                                    current_f_si = Some(
                                        String::from_utf8_lossy(&attr.value).to_string()
                                    );
                                }
                                b"ref" => {
                                    current_f_has_ref = true;
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_formula => {
                let text = String::from_utf8_lossy(e.as_ref()).to_string();
                if !text.is_empty() {
                    current_f_text = Some(text);
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"f" => {
                        // End of <f>...</f> — record the entry
                        if let Some(ref cell_ref) = current_cell_ref {
                            if let Some((row, col)) = parse_xlsx_cell_ref(cell_ref) {
                                let is_master = current_f_shared && current_f_has_ref && current_f_text.is_some();
                                if is_master {
                                    // Shared formula master: store definition
                                    if let Some(ref si) = current_f_si {
                                        shared_defs.insert(si.clone(), SharedFormulaDef {
                                            base_row: row,
                                            base_col: col,
                                            formula: current_f_text.clone().unwrap(),
                                        });
                                    }
                                }
                                entries.push(FormulaEntry {
                                    row,
                                    col,
                                    formula_text: current_f_text.take(),
                                    shared_si: current_f_si.take(),
                                });
                            }
                        }
                        in_formula = false;
                    }
                    b"c" => {
                        current_cell_ref = None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if e.name().as_ref() == b"f" => {
                // <f t="shared" si="N"/> — shared formula follower (empty element)
                let mut si = None;
                let mut is_shared = false;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"t" if attr.value.as_ref() == b"shared" => {
                            is_shared = true;
                        }
                        b"si" => {
                            si = Some(String::from_utf8_lossy(&attr.value).to_string());
                        }
                        _ => {}
                    }
                }
                if is_shared {
                    if let Some(ref cell_ref) = current_cell_ref {
                        if let Some((row, col)) = parse_xlsx_cell_ref(cell_ref) {
                            entries.push(FormulaEntry {
                                row,
                                col,
                                formula_text: None, // follower — no text
                                shared_si: si,
                            });
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // --- Pass 2: Resolve followers using shared formula definitions ---
    for entry in &entries {
        let formula = if let Some(ref text) = entry.formula_text {
            // Master or normal formula — use as-is
            text.clone()
        } else if let Some(ref si) = entry.shared_si {
            // Follower — look up master and shift references
            if let Some(def) = shared_defs.get(si) {
                let row_delta = entry.row as i32 - def.base_row as i32;
                let col_delta = entry.col as i32 - def.base_col as i32;
                if row_delta == 0 && col_delta == 0 {
                    def.formula.clone()
                } else {
                    adjust_formula_refs_for_shared(&def.formula, row_delta, col_delta)
                }
            } else {
                continue; // No master found — skip
            }
        } else {
            continue; // No formula text and not shared — skip
        };

        out.push((sheet_idx, entry.row, entry.col, formula));
    }
}

/// Adjust cell references in a formula by row/col deltas, respecting $ anchors.
/// Used for expanding shared formula followers from their master definition.
fn adjust_formula_refs_for_shared(formula: &str, row_delta: i32, col_delta: i32) -> String {
    use regex::Regex;

    // Match cell references: optional $ before col letters, col letters, optional $ before row, row digits
    // Examples: A1, $A$1, A$1, $A1, AA100, R103
    let re = Regex::new(r"(\$?)([A-Za-z]+)(\$?)(\d+)").unwrap();

    re.replace_all(formula, |caps: &regex::Captures| {
        let col_absolute = &caps[1] == "$";
        let col_letters = &caps[2];
        let row_absolute = &caps[3] == "$";
        let row_num: i32 = caps[4].parse().unwrap_or(1);

        // Don't adjust function names that look like cell refs (they won't have digits though)
        // The regex requires trailing digits, so function names like SUM, IF won't match.

        // Parse column letters to 0-indexed number
        let col = col_letters.to_uppercase().chars().fold(0i32, |acc, c| {
            acc * 26 + (c as i32 - 'A' as i32 + 1)
        }) - 1;

        // Apply deltas (skip if absolute)
        let new_col = if col_absolute { col } else { col + col_delta };
        let new_row = if row_absolute { row_num } else { row_num + row_delta };

        // Bounds check
        if new_col < 0 || new_row < 1 {
            return "#REF!".to_string();
        }

        // Convert column back to letters
        let col_str = col_num_to_letters(new_col as usize);

        format!(
            "{}{}{}{}",
            if col_absolute { "$" } else { "" },
            col_str,
            if row_absolute { "$" } else { "" },
            new_row
        )
    })
    .to_string()
}

/// Convert a 0-indexed column number to Excel column letters (0=A, 1=B, ..., 25=Z, 26=AA)
fn col_num_to_letters(mut col: usize) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result
}

/// Extract value-only cells from XLSX XML that calamine may have missed.
/// Parses shared strings (t="s"), inline strings (t="inlineStr"), and numeric values.
/// Returns: Vec<(sheet_index, row, col, value_string)>
fn extract_xml_value_cells(path: &Path) -> Vec<(usize, usize, usize, String)> {
    use zip::ZipArchive;

    let file = match std::fs::File::open(path) {
        Ok(f) => f,
        Err(_) => return Vec::new(),
    };
    let mut archive = match ZipArchive::new(file) {
        Ok(a) => a,
        Err(_) => return Vec::new(),
    };

    // Load shared strings table
    let shared_strings = match read_zip_file_for_shared(&mut archive, "xl/sharedStrings.xml") {
        Some(xml) => parse_shared_strings(&xml),
        None => Vec::new(),
    };

    let workbook_xml = match read_zip_file_for_shared(&mut archive, "xl/workbook.xml") {
        Some(s) => s,
        None => return Vec::new(),
    };
    let rels_xml = match read_zip_file_for_shared(&mut archive, "xl/_rels/workbook.xml.rels") {
        Some(s) => s,
        None => return Vec::new(),
    };

    let worksheet_paths = resolve_worksheet_paths(&workbook_xml, &rels_xml);

    let mut all_values = Vec::new();

    for (sheet_idx, ws_path) in worksheet_paths.iter().enumerate() {
        let xml = match read_zip_file_for_shared(&mut archive, ws_path) {
            Some(s) => s,
            None => continue,
        };
        extract_values_from_xml(&xml, sheet_idx, &shared_strings, &mut all_values);
    }

    all_values
}

/// Parse xl/sharedStrings.xml into a Vec of strings indexed by position.
/// Each <si> element contains either <t>text</t> or <r><t>text</t></r> (rich text).
fn parse_shared_strings(xml: &str) -> Vec<String> {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut strings = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false); // preserve whitespace in string values
    let mut buf = Vec::new();

    let mut in_si = false;
    let mut in_t = false;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"si" => {
                        in_si = true;
                        current_text.clear();
                    }
                    b"t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_t => {
                current_text.push_str(&String::from_utf8_lossy(e.as_ref()));
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"t" => {
                        in_t = false;
                    }
                    b"si" => {
                        strings.push(current_text.clone());
                        in_si = false;
                        current_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    strings
}

/// Parse a single worksheet XML and extract value-only cells (no formula).
/// Handles cell types: shared string (t="s"), inline string (t="inlineStr"),
/// numeric (t="n" or no type), boolean (t="b"), string (t="str").
fn extract_values_from_xml(
    xml: &str,
    sheet_idx: usize,
    shared_strings: &[String],
    out: &mut Vec<(usize, usize, usize, String)>,
) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut current_cell_ref: Option<String> = None;
    let mut current_cell_type: Option<String> = None; // t attribute
    let mut has_formula = false;
    let mut in_value = false;     // inside <v>
    let mut in_inline_t = false;  // inside <is><t>
    let mut in_is = false;        // inside <is>
    let mut value_text = String::new();
    let mut inline_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"c" => {
                        current_cell_ref = None;
                        current_cell_type = None;
                        has_formula = false;
                        value_text.clear();
                        inline_text.clear();
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    current_cell_ref = Some(
                                        String::from_utf8_lossy(&attr.value).to_string()
                                    );
                                }
                                b"t" => {
                                    current_cell_type = Some(
                                        String::from_utf8_lossy(&attr.value).to_string()
                                    );
                                }
                                _ => {}
                            }
                        }
                    }
                    b"f" => {
                        has_formula = true;
                    }
                    b"v" => {
                        in_value = true;
                        value_text.clear();
                    }
                    b"is" => {
                        in_is = true;
                    }
                    b"t" if in_is => {
                        in_inline_t = true;
                        inline_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_value {
                    value_text.push_str(&String::from_utf8_lossy(e.as_ref()));
                } else if in_inline_t {
                    inline_text.push_str(&String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::End(ref e)) => {
                match e.name().as_ref() {
                    b"v" => {
                        in_value = false;
                    }
                    b"t" if in_is => {
                        in_inline_t = false;
                    }
                    b"is" => {
                        in_is = false;
                    }
                    b"c" => {
                        // End of cell — emit value if no formula and has content
                        if !has_formula {
                            if let Some(ref cell_ref) = current_cell_ref {
                                if let Some((row, col)) = parse_xlsx_cell_ref(cell_ref) {
                                    let resolved = resolve_cell_value(
                                        current_cell_type.as_deref(),
                                        &value_text,
                                        &inline_text,
                                        shared_strings,
                                    );
                                    if let Some(val) = resolved {
                                        out.push((sheet_idx, row, col, val));
                                    }
                                }
                            }
                        }
                        current_cell_ref = None;
                        current_cell_type = None;
                        has_formula = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                match e.name().as_ref() {
                    b"c" => {
                        // Empty cell element <c r="..." .../> — no value, skip
                    }
                    b"f" => {
                        has_formula = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

/// Resolve a cell value from its XLSX type and raw text.
/// Returns None for empty/unresolvable cells.
fn resolve_cell_value(
    cell_type: Option<&str>,
    value_text: &str,
    inline_text: &str,
    shared_strings: &[String],
) -> Option<String> {
    match cell_type {
        Some("s") => {
            // Shared string: value_text is the index into shared strings table
            let idx: usize = value_text.parse().ok()?;
            shared_strings.get(idx).cloned()
        }
        Some("inlineStr") => {
            // Inline string: text is in <is><t>...</t></is>
            if inline_text.is_empty() {
                None
            } else {
                Some(inline_text.to_string())
            }
        }
        Some("b") => {
            // Boolean: 0 = FALSE, 1 = TRUE
            match value_text {
                "1" => Some("TRUE".to_string()),
                "0" => Some("FALSE".to_string()),
                _ => None,
            }
        }
        Some("str") => {
            // Cached string result of a formula — but we only emit non-formula cells
            if value_text.is_empty() {
                None
            } else {
                Some(value_text.to_string())
            }
        }
        Some("n") | None => {
            // Numeric (explicit or default): value_text is the number
            if value_text.is_empty() {
                None
            } else {
                // Clean up integer representation: "0.1" stays, "5" stays
                Some(value_text.to_string())
            }
        }
        Some("e") => {
            // Error value
            if value_text.is_empty() {
                None
            } else {
                Some(value_text.to_string())
            }
        }
        Some(_) => {
            // Unknown type — try to preserve the value
            if value_text.is_empty() {
                None
            } else {
                Some(value_text.to_string())
            }
        }
    }
}

/// Parse an XLSX cell reference like "R104" or "AA1" to (row, col) in 0-indexed.
fn parse_xlsx_cell_ref(cell_ref: &str) -> Option<(usize, usize)> {
    let mut col_part = String::new();
    let mut row_part = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col_part.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            row_part.push(ch);
        }
    }

    if col_part.is_empty() || row_part.is_empty() {
        return None;
    }

    // Convert column letters to 0-indexed number (A=0, B=1, ..., Z=25, AA=26)
    let mut col: usize = 0;
    for ch in col_part.chars() {
        col = col * 26 + (ch as usize - 'A' as usize + 1);
    }
    col -= 1; // 0-indexed

    // Convert row string to 0-indexed number
    let row: usize = row_part.parse::<usize>().ok()? - 1;

    Some((row, col))
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

    // Font size
    if let Some(size) = cell_format.font_size {
        format = format.set_font_size(size as f64);
    }

    // Font color
    if let Some([r, g, b, _]) = cell_format.font_color {
        let color = rust_xlsxwriter::Color::RGB(
            ((r as u32) << 16) | ((g as u32) << 8) | (b as u32)
        );
        format = format.set_font_color(color);
    }

    // Font family
    if let Some(ref family) = cell_format.font_family {
        format = format.set_font_name(family);
    }

    // Horizontal alignment
    format = match cell_format.alignment {
        Alignment::General => format, // Excel default: numbers right, text left
        Alignment::Left => format.set_align(FormatAlign::Left),
        Alignment::Center => format.set_align(FormatAlign::Center),
        Alignment::Right => format.set_align(FormatAlign::Right),
        Alignment::CenterAcrossSelection => format.set_align(FormatAlign::CenterAcross),
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

/// Escape a currency symbol for use in Excel format codes.
/// Simple symbols ($, €, £, ¥) pass through; others get quoted.
fn excel_literal_prefix(sym: &str) -> String {
    if sym.chars().all(|c| c.is_ascii_alphanumeric() || "$€£¥".contains(c)) {
        sym.to_string()
    } else {
        format!("\"{}\"", sym.replace('"', "\"\""))
    }
}

/// Build the numeric pattern portion of a format code.
/// e.g. `#,##0.00` (thousands=true, decimals=2) or `0.00` (thousands=false, decimals=2)
fn build_number_pattern(decimals: u8, thousands: bool) -> String {
    let int_part = if thousands { "#,##0" } else { "0" };
    if decimals == 0 {
        int_part.to_string()
    } else {
        format!("{}.{}", int_part, "0".repeat(decimals as usize))
    }
}

/// Apply number format to an Excel Format
fn apply_number_format(format: Format, number_format: &NumberFormat) -> Format {
    match number_format {
        NumberFormat::General => format,
        NumberFormat::Number { decimals, thousands, negative } => {
            let pos = build_number_pattern(*decimals, *thousands);
            let neg = match negative {
                visigrid_engine::cell::NegativeStyle::Minus => format!("-{}", pos),
                visigrid_engine::cell::NegativeStyle::Parens => format!("({})", pos),
                visigrid_engine::cell::NegativeStyle::RedMinus => format!("[Red]-{}", pos),
                visigrid_engine::cell::NegativeStyle::RedParens => format!("[Red]({})", pos),
            };
            let pattern = format!("{};{};{};@", pos, neg, pos);
            format.set_num_format(&pattern)
        }
        NumberFormat::Currency { decimals, thousands, negative, symbol } => {
            let sym = excel_literal_prefix(symbol.as_deref().unwrap_or("$"));
            let num_pat = build_number_pattern(*decimals, *thousands);
            let pos = format!("{}{}", sym, num_pat);
            let neg = match negative {
                visigrid_engine::cell::NegativeStyle::Minus => format!("-{}", pos),
                visigrid_engine::cell::NegativeStyle::Parens => format!("({})", pos),
                visigrid_engine::cell::NegativeStyle::RedMinus => format!("[Red]-{}", pos),
                visigrid_engine::cell::NegativeStyle::RedParens => format!("[Red]({})", pos),
            };
            let pattern = format!("{};{};{};@", pos, neg, pos);
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
        NumberFormat::Custom(code) => format.set_num_format(code),
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
        || format.font_size.is_some()
        || format.font_color.is_some()
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
    fn test_strip_ods_prefix_basic() {
        assert_eq!(strip_ods_prefix("=of:SUM(A1:A10)"), "=SUM(A1:A10)");
        assert_eq!(strip_ods_prefix("=OF:SUM(A1:A10)"), "=SUM(A1:A10)");
    }

    #[test]
    fn test_strip_ods_prefix_nested_and_semicolons() {
        // ODS uses semicolons as argument separators — should become commas
        assert_eq!(
            strip_ods_prefix("=of:IF(of:AND(A1>0;B1>0);1;0)"),
            "=IF(AND(A1>0,B1>0),1,0)"
        );
        assert_eq!(
            strip_ods_prefix("=of:VLOOKUP(A1;B1:C10;2;0)"),
            "=VLOOKUP(A1,B1:C10,2,0)"
        );
    }

    #[test]
    fn test_strip_ods_prefix_preserves_strings() {
        // Semicolons inside string literals must NOT be converted
        assert_eq!(
            strip_ods_prefix("=of:IF(A1>0;\"yes;no\";\"maybe\")"),
            "=IF(A1>0,\"yes;no\",\"maybe\")"
        );
    }

    #[test]
    fn test_strip_ods_prefix_passthrough() {
        // Non-ODS formulas should pass through unchanged
        assert_eq!(strip_ods_prefix("=SUM(A1:A10)"), "=SUM(A1:A10)");
        assert_eq!(strip_ods_prefix("=A1+B1"), "=A1+B1");
    }

    #[test]
    fn test_import_result_summary() {
        let mut result = ImportResult::default();
        result.sheets_imported = 1;
        result.cells_imported = 100;
        result.formulas_imported = 0;

        assert_eq!(result.summary(), "1 sheet · 100 cells");

        result.sheets_imported = 3;
        result.formulas_imported = 25;
        assert_eq!(result.summary(), "3 sheets · 100 cells · 25 formulas");
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
        format.number_format = NumberFormat::currency_compat(2);
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
    // Merge export tests
    // ========================================================================

    #[test]
    fn test_xlsx_export_merged_cells_no_leak() {
        use visigrid_engine::sheet::MergedRegion;

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut();

        // Set values that will become merge-hidden
        sheet.set_value(0, 0, "Header");
        sheet.set_value(0, 1, "LEAK1");
        sheet.set_value(0, 2, "LEAK2");
        sheet.set_value(1, 0, "10");
        sheet.set_value(1, 1, "20");
        sheet.set_value(1, 2, "30");

        // Merge A1:C1 — B1/C1 become hidden but hold residual data
        sheet.add_merge(MergedRegion::new(0, 0, 0, 2)).unwrap();

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("test_merge_leak.xlsx");

        let result = export(&workbook, &export_path, None).unwrap();
        assert_eq!(result.merges_exported, 1);

        // Re-import and verify origin value survives, hidden values don't leak
        let (imported_wb, import_result) = import(&export_path).expect("Import should succeed");
        let imported = &imported_wb.sheets()[0];

        // Origin cell must retain its value
        assert_eq!(imported.get_display(0, 0), "Header", "merge origin should survive roundtrip");

        // Hidden cells must not contain residual data
        let b1 = imported.get_display(0, 1);
        let c1 = imported.get_display(0, 2);
        assert!(b1.is_empty(), "hidden merge cell B1 leaked: {b1}");
        assert!(c1.is_empty(), "hidden merge cell C1 leaked: {c1}");

        // Non-merged cells are unaffected
        assert_eq!(imported.get_display(1, 0), "10");
        assert_eq!(imported.get_display(1, 1), "20");
        assert_eq!(imported.get_display(1, 2), "30");

        // Merge structure survived roundtrip
        assert_eq!(import_result.merges_imported, 1);
        assert!(imported.is_merge_hidden(0, 1), "B1 should be merge-hidden after import");
        assert!(imported.is_merge_hidden(0, 2), "C1 should be merge-hidden after import");
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
        format.number_format = NumberFormat::currency_compat(0);
        sheet.set_format(0, 0, format);

        // Currency with 2 decimals
        sheet.set_value(1, 0, "1234.56");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::currency_compat(2);
        sheet.set_format(1, 0, format);

        // Negative currency
        sheet.set_value(2, 0, "-1234.56");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::currency_compat(2);
        sheet.set_format(2, 0, format);

        // Zero currency
        sheet.set_value(3, 0, "0");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::currency_compat(2);
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
        format.number_format = NumberFormat::number_compat(7);
        sheet.set_format(1, 0, format);

        // Negative with many decimals
        sheet.set_value(2, 0, "-123.456789");
        let mut format = CellFormat::default();
        format.number_format = NumberFormat::number_compat(6);
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
            price_fmt.number_format = NumberFormat::currency_compat(2);
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
        total_fmt.number_format = NumberFormat::currency_compat(2);
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
        curr_fmt.number_format = NumberFormat::currency_compat(2);
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

    // ========================================================================
    // Per-sheet layout export tests
    //
    // NOTE: These are export-only tests. Calamine (our import library) doesn't
    // support reading column widths or row heights - it's focused on cell data.
    // We verify export correctness by parsing the XLSX XML directly.
    //
    // TODO: Implement layout import by parsing XLSX XML ourselves (like we do
    // in these tests), then add full round-trip tests. This would allow:
    // - Preserving user's column widths when re-importing an exported file
    // - Importing Excel files with custom layouts
    // ========================================================================

    #[test]
    fn test_xlsx_export_writes_per_sheet_column_widths() {
        // This test verifies that column widths are correctly written per-sheet
        // when exporting to XLSX. It's a critical regression test for the
        // per-sheet sizing feature.
        //
        // We parse the XLSX ZIP directly to verify the column widths in each
        // sheet's XML, since calamine doesn't support reading layout data.

        use std::io::Read as IoRead;

        let mut workbook = Workbook::new();

        // Sheet 1: set data and prepare custom widths
        workbook.active_sheet_mut().set_value(0, 0, "Sheet1 Data");
        workbook.active_sheet_mut().set_value(0, 1, "Col B");
        workbook.active_sheet_mut().set_value(0, 2, "Col C");

        // Sheet 2: add and set data
        workbook.add_sheet();
        workbook.sheet_mut(1).unwrap().set_value(0, 0, "Sheet2 Data");
        workbook.sheet_mut(1).unwrap().set_value(0, 1, "Different Layout");

        // Sheet 3: add and set data
        workbook.add_sheet();
        workbook.sheet_mut(2).unwrap().set_value(0, 0, "Sheet3 Data");

        // Create per-sheet layouts with DIFFERENT column widths
        // Sheet 1: wide column A (200px), normal B, narrow C
        let mut layout1 = ExportLayout::default();
        layout1.col_widths.insert(0, 200.0);  // Wide
        layout1.col_widths.insert(1, 100.0);  // Normal
        layout1.col_widths.insert(2, 50.0);   // Narrow

        // Sheet 2: narrow column A, wide B
        let mut layout2 = ExportLayout::default();
        layout2.col_widths.insert(0, 70.0);   // Narrow
        layout2.col_widths.insert(1, 250.0);  // Very wide

        // Sheet 3: only default widths (no customization)
        let layout3 = ExportLayout::default();

        let layouts = vec![layout1, layout2, layout3];

        // Export
        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("per_sheet_widths.xlsx");

        let result = export(&workbook, &export_path, Some(&layouts)).unwrap();
        assert_eq!(result.sheets_exported, 3);

        // Parse the XLSX to verify per-sheet column widths
        let file = std::fs::File::open(&export_path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        // Helper to extract col widths from worksheet XML
        // XLSX XML is typically minified (no newlines), so we parse it character by character
        fn extract_col_widths(xml: &str) -> Vec<(usize, f64)> {
            let mut widths = Vec::new();
            let mut search_start = 0;

            // Find all <col ... /> or <col ...></col> elements
            while let Some(col_start) = xml[search_start..].find("<col ") {
                let abs_start = search_start + col_start;
                // Find the end of this tag (either /> or >)
                let tag_end = xml[abs_start..].find("/>")
                    .or_else(|| xml[abs_start..].find('>'))
                    .map(|e| abs_start + e)
                    .unwrap_or(xml.len());

                let col_tag = &xml[abs_start..tag_end];

                // Helper to extract attribute value
                fn get_attr(tag: &str, attr: &str) -> Option<String> {
                    let search = format!("{}=\"", attr);
                    if let Some(pos) = tag.find(&search) {
                        let start = pos + search.len();
                        if let Some(end) = tag[start..].find('"') {
                            return Some(tag[start..start + end].to_string());
                        }
                    }
                    None
                }

                // Extract min, max, and width attributes
                if let (Some(min_str), Some(width_str)) = (get_attr(col_tag, "min"), get_attr(col_tag, "width")) {
                    if let (Ok(min), Ok(width)) = (min_str.parse::<usize>(), width_str.parse::<f64>()) {
                        // max defaults to min if not specified
                        let max = get_attr(col_tag, "max")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(min);

                        // Add each column in the range (convert to 0-based)
                        for col in min..=max {
                            widths.push((col - 1, width));
                        }
                    }
                }

                search_start = tag_end;
            }
            widths
        }

        // Read sheet1.xml
        let mut sheet1_xml = String::new();
        archive.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut sheet1_xml).unwrap();
        let sheet1_widths = extract_col_widths(&sheet1_xml);

        // Read sheet2.xml
        let mut sheet2_xml = String::new();
        archive.by_name("xl/worksheets/sheet2.xml").unwrap().read_to_string(&mut sheet2_xml).unwrap();
        let sheet2_widths = extract_col_widths(&sheet2_xml);

        // Read sheet3.xml
        let mut sheet3_xml = String::new();
        archive.by_name("xl/worksheets/sheet3.xml").unwrap().read_to_string(&mut sheet3_xml).unwrap();
        let sheet3_widths = extract_col_widths(&sheet3_xml);

        // Verify Sheet 1 has 3 custom column widths
        assert_eq!(sheet1_widths.len(), 3, "Sheet1 should have 3 custom column widths");

        // Verify Sheet 2 has 2 custom column widths
        assert_eq!(sheet2_widths.len(), 2, "Sheet2 should have 2 custom column widths");

        // Verify Sheet 3 has no custom column widths (uses defaults)
        assert_eq!(sheet3_widths.len(), 0, "Sheet3 should have no custom column widths");

        // Verify the actual widths are different between sheets
        // Sheet 1 col A should be wider than Sheet 2 col A (200px vs 70px)
        let sheet1_col_a_width = sheet1_widths.iter().find(|(col, _)| *col == 0).map(|(_, w)| *w).unwrap_or(0.0);
        let sheet2_col_a_width = sheet2_widths.iter().find(|(col, _)| *col == 0).map(|(_, w)| *w).unwrap_or(0.0);

        // 200px / 7 ≈ 28.6 Excel units, 70px / 7 = 10 Excel units
        assert!(sheet1_col_a_width > sheet2_col_a_width,
            "Sheet1 col A ({:.1}) should be wider than Sheet2 col A ({:.1})",
            sheet1_col_a_width, sheet2_col_a_width);

        // Sheet 2 col B should be very wide (250px → ~35.7 Excel units)
        let sheet2_col_b_width = sheet2_widths.iter().find(|(col, _)| *col == 1).map(|(_, w)| *w).unwrap_or(0.0);
        assert!(sheet2_col_b_width > 30.0,
            "Sheet2 col B ({:.1}) should be very wide (>30 Excel units)",
            sheet2_col_b_width);

        // Re-import and verify data integrity (widths not preserved in import,
        // but data should be intact)
        let (reimported, _) = import(&export_path).unwrap();
        assert_eq!(reimported.sheet_count(), 3);
        assert_eq!(reimported.sheet(0).unwrap().get_display(0, 0), "Sheet1 Data");
        assert_eq!(reimported.sheet(1).unwrap().get_display(0, 0), "Sheet2 Data");
        assert_eq!(reimported.sheet(2).unwrap().get_display(0, 0), "Sheet3 Data");
    }

    #[test]
    fn test_xlsx_export_writes_per_sheet_row_heights() {
        // Verify row heights are written per-sheet correctly to XLSX.
        // We parse the XML directly since calamine doesn't read layout data.

        use std::io::Read as IoRead;

        let mut workbook = Workbook::new();
        workbook.active_sheet_mut().set_value(0, 0, "Sheet1 Row 0");
        workbook.add_sheet();
        workbook.sheet_mut(1).unwrap().set_value(0, 0, "Sheet2 Row 0");
        workbook.sheet_mut(1).unwrap().set_value(1, 0, "Sheet2 Row 1");

        // Sheet 1: very tall row 0 (100px → ~75 points)
        let mut layout1 = ExportLayout::default();
        layout1.row_heights.insert(0, 100.0);

        // Sheet 2: medium tall row 0 (40px → ~30 points) - distinct from default (~15pt)
        let mut layout2 = ExportLayout::default();
        layout2.row_heights.insert(0, 40.0);

        let layouts = vec![layout1, layout2];

        let temp_dir = tempfile::tempdir().unwrap();
        let export_path = temp_dir.path().join("per_sheet_row_heights.xlsx");

        let result = export(&workbook, &export_path, Some(&layouts)).unwrap();
        assert_eq!(result.sheets_exported, 2);

        // Parse the XLSX to verify per-sheet row heights differ
        let file = std::fs::File::open(&export_path).unwrap();
        let mut archive = zip::ZipArchive::new(file).unwrap();

        // Helper to extract row 1's height from worksheet XML
        fn extract_row_height(xml: &str, row_num: usize) -> Option<f64> {
            // Look for <row r="N" ... ht="H" ...>
            let row_search = format!("r=\"{}\"", row_num);
            let mut search_start = 0;

            while let Some(row_start) = xml[search_start..].find("<row ") {
                let abs_start = search_start + row_start;
                let tag_end = xml[abs_start..].find('>').map(|e| abs_start + e).unwrap_or(xml.len());
                let row_tag = &xml[abs_start..tag_end];

                if row_tag.contains(&row_search) {
                    // Extract ht attribute
                    if let Some(ht_pos) = row_tag.find("ht=\"") {
                        let ht_start = ht_pos + 4;
                        if let Some(ht_end) = row_tag[ht_start..].find('"') {
                            if let Ok(height) = row_tag[ht_start..ht_start + ht_end].parse::<f64>() {
                                return Some(height);
                            }
                        }
                    }
                }
                search_start = tag_end;
            }
            None
        }

        let mut sheet1_xml = String::new();
        archive.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut sheet1_xml).unwrap();

        let mut sheet2_xml = String::new();
        archive.by_name("xl/worksheets/sheet2.xml").unwrap().read_to_string(&mut sheet2_xml).unwrap();

        // Get row 1 heights from each sheet (1-based in XML, so row index 0 = r="1")
        let sheet1_row0_height = extract_row_height(&sheet1_xml, 1);
        let sheet2_row0_height = extract_row_height(&sheet2_xml, 1);

        // Both sheets should have row 0 (r="1") height defined
        assert!(sheet1_row0_height.is_some(), "Sheet1 row 0 should have custom height");
        assert!(sheet2_row0_height.is_some(), "Sheet2 row 0 should have custom height");

        // Sheet 1's row 0 should be much taller than Sheet 2's row 0
        // (100px / 1.33 ≈ 75pt vs 20px / 1.33 ≈ 15pt)
        let h1 = sheet1_row0_height.unwrap();
        let h2 = sheet2_row0_height.unwrap();
        assert!(h1 > h2 * 2.0,
            "Sheet1 row 0 height ({:.1}) should be much taller than Sheet2 row 0 ({:.1})",
            h1, h2);
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

    /// Test that import detects recalc errors and circular references
    #[test]
    fn test_import_recalc_error_counting() {
        let mut result = ImportResult::default();
        result.recalc_errors = 5;
        result.recalc_circular = 2;

        assert!(result.has_warnings());
        let summary = result.warning_summary().unwrap();
        assert!(summary.contains("5 formula errors"), "summary: {}", summary);
        assert!(summary.contains("2 circular references"), "summary: {}", summary);
    }

    /// Test that import with no recalc errors reports clean
    #[test]
    fn test_import_no_recalc_errors() {
        let result = ImportResult::default();
        assert!(!result.has_warnings());
        assert!(result.warning_summary().is_none());
    }

    /// Test shared formula XML detection
    #[test]
    fn test_count_shared_masters_in_xml() {
        // Worksheet XML with 2 shared formula master groups
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet>
            <sheetData>
                <row r="1">
                    <c r="A1"><v>10</v></c>
                    <c r="B1"><f>A1*2</f><v>20</v></c>
                </row>
                <row r="2">
                    <c r="A2"><v>20</v></c>
                    <c r="B2"><f t="shared" ref="B2:B10" si="0">A2*2</f><v>40</v></c>
                </row>
                <row r="3">
                    <c r="A3"><v>30</v></c>
                    <c r="B3"><f t="shared" si="0"/><v>60</v></c>
                </row>
                <row r="4">
                    <c r="C4"><f t="shared" ref="C4:C10" si="1">SUM(A4:B4)</f><v>0</v></c>
                </row>
                <row r="5">
                    <c r="C5"><f t="shared" si="1"/><v>0</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        // Should count 2 masters (B2 with ref="B2:B10" and C4 with ref="C4:C10")
        // B3 and C5 are dependents (no ref= attribute), so they don't count
        let count = count_shared_masters_in_xml(xml);
        assert_eq!(count, 2, "Should detect 2 shared formula master groups");
    }

    /// Test shared formula detection on a regular formula (no shared)
    #[test]
    fn test_count_shared_masters_no_shared() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet>
            <sheetData>
                <row r="1">
                    <c r="A1"><f>SUM(B1:B10)</f><v>0</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let count = count_shared_masters_in_xml(xml);
        assert_eq!(count, 0, "Regular formulas should not count as shared");
    }

    /// Test that a simple XLSX roundtrip has zero recalc errors
    #[test]
    fn test_simple_xlsx_roundtrip_no_errors() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsx");

        // Create a workbook with some formulas
        let mut wb = Workbook::new();
        let sheet = wb.sheet_mut(0).unwrap();
        sheet.set_value(0, 0, "10");
        sheet.set_value(1, 0, "20");
        sheet.set_value(2, 0, "30");
        sheet.set_value(0, 1, "=A1*2");
        sheet.set_value(1, 1, "=A2*2");
        sheet.set_value(2, 1, "=A3*2");
        sheet.set_value(3, 0, "=SUM(A1:A3)");

        // Export to XLSX
        let export_result = export(&wb, &path, None).expect("Export should succeed");
        assert!(export_result.formulas_exported > 0);

        // Import back
        let (imported_wb, result) = import(&path).expect("Import should succeed");

        // No recalc errors
        assert_eq!(result.recalc_errors, 0, "Should have 0 recalc errors");
        assert_eq!(result.recalc_circular, 0, "Should have 0 circular refs");

        // Verify some sentinel values
        let sheet = &imported_wb.sheets()[0];
        assert_eq!(sheet.get_display(0, 1), "20", "A1*2 should be 20");
        assert_eq!(sheet.get_display(1, 1), "40", "A2*2 should be 40");
        assert_eq!(sheet.get_display(3, 0), "60", "SUM(A1:A3) should be 60");
    }

    /// Test XLSX roundtrip with relative references and range copy-down.
    ///
    /// This exercises the patterns that shared formula expansion must handle:
    /// relative cell refs that shift when copied down, and range references
    /// that shift row-by-row. These are the patterns calamine 0.26 broke.
    ///
    /// Note: VisiGrid's formula parser does not yet support $ (absolute)
    /// reference anchors. Formulas with $A$1, A$1, $A1 are tracked as parse
    /// errors at import time. Absolute anchor support is a separate engine task.
    #[test]
    fn test_xlsx_roundtrip_relative_references() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("rel_refs.xlsx");

        let mut wb = Workbook::new();
        let sheet = wb.sheet_mut(0).unwrap();

        // Data column A: values 10, 20, 30, 40
        sheet.set_value(0, 0, "10");  // A1
        sheet.set_value(1, 0, "20");  // A2
        sheet.set_value(2, 0, "30");  // A3
        sheet.set_value(3, 0, "40");  // A4

        // Data column B: values 1, 2, 3, 4
        sheet.set_value(0, 1, "1");   // B1
        sheet.set_value(1, 1, "2");   // B2
        sheet.set_value(2, 1, "3");   // B3
        sheet.set_value(3, 1, "4");   // B4

        // Column C: relative ref copy-down (=A1*B1, =A2*B2, =A3*B3)
        // This is the pattern shared formulas produce when expanded
        sheet.set_value(0, 2, "=A1*B1");  // C1 = 10
        sheet.set_value(1, 2, "=A2*B2");  // C2 = 40
        sheet.set_value(2, 2, "=A3*B3");  // C3 = 90
        sheet.set_value(3, 2, "=A4*B4");  // C4 = 160

        // Column D: range SUM copied down (=SUM(A1:B1), =SUM(A2:B2), ...)
        sheet.set_value(0, 3, "=SUM(A1:B1)");  // D1 = 11
        sheet.set_value(1, 3, "=SUM(A2:B2)");  // D2 = 22
        sheet.set_value(2, 3, "=SUM(A3:B3)");  // D3 = 33
        sheet.set_value(3, 3, "=SUM(A4:B4)");  // D4 = 44

        // Column E: cross-column ref (=C1+D1, =C2+D2, ...)
        sheet.set_value(0, 4, "=C1+D1");  // E1 = 21
        sheet.set_value(1, 4, "=C2+D2");  // E2 = 62
        sheet.set_value(2, 4, "=C3+D3");  // E3 = 123

        // Column F: nested function with relative range
        sheet.set_value(0, 5, "=AVERAGE(A1:A4)");  // F1 = 25
        sheet.set_value(1, 5, "=MAX(B1:B4)");      // F2 = 4
        sheet.set_value(2, 5, "=MIN(A1:A4)");      // F3 = 10

        // Export to XLSX
        let export_result = export(&wb, &path, None).expect("Export should succeed");
        assert!(export_result.formulas_exported >= 13, "Should export all formulas");

        // Import back
        let (imported_wb, result) = import(&path).expect("Import should succeed");

        // Zero recalc errors — the critical assertion
        assert_eq!(result.recalc_errors, 0,
            "Relative refs should produce 0 errors, got {}. Examples: {:?}",
            result.recalc_errors,
            result.recalc_error_examples.iter().map(|e| format!("{}!{}: {}", e.sheet, e.address, e.error)).collect::<Vec<_>>()
        );
        assert_eq!(result.recalc_circular, 0, "Relative refs should produce 0 circulars");

        let sheet = &imported_wb.sheets()[0];

        // Verify relative multiplication copy-down
        assert_eq!(sheet.get_display(0, 2), "10", "C1: =A1*B1 should be 10");
        assert_eq!(sheet.get_display(1, 2), "40", "C2: =A2*B2 should be 40");
        assert_eq!(sheet.get_display(2, 2), "90", "C3: =A3*B3 should be 90");
        assert_eq!(sheet.get_display(3, 2), "160", "C4: =A4*B4 should be 160");

        // Verify range SUM copy-down
        assert_eq!(sheet.get_display(0, 3), "11", "D1: =SUM(A1:B1) should be 11");
        assert_eq!(sheet.get_display(1, 3), "22", "D2: =SUM(A2:B2) should be 22");
        assert_eq!(sheet.get_display(2, 3), "33", "D3: =SUM(A3:B3) should be 33");
        assert_eq!(sheet.get_display(3, 3), "44", "D4: =SUM(A4:B4) should be 44");

        // Verify cross-column references
        assert_eq!(sheet.get_display(0, 4), "21", "E1: =C1+D1 should be 21");
        assert_eq!(sheet.get_display(1, 4), "62", "E2: =C2+D2 should be 62");
        assert_eq!(sheet.get_display(2, 4), "123", "E3: =C3+D3 should be 123");

        // Verify nested functions
        assert_eq!(sheet.get_display(0, 5), "25", "F1: =AVERAGE(A1:A4) should be 25");
        assert_eq!(sheet.get_display(1, 5), "4", "F2: =MAX(B1:B4) should be 4");
        assert_eq!(sheet.get_display(2, 5), "10", "F3: =MIN(A1:A4) should be 10");
    }

    #[test]
    fn test_parse_xlsx_cell_ref() {
        assert_eq!(parse_xlsx_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_xlsx_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_xlsx_cell_ref("Z1"), Some((0, 25)));
        assert_eq!(parse_xlsx_cell_ref("AA1"), Some((0, 26)));
        assert_eq!(parse_xlsx_cell_ref("R104"), Some((103, 17)));
        assert_eq!(parse_xlsx_cell_ref("IV256"), Some((255, 255)));
        assert_eq!(parse_xlsx_cell_ref(""), None);
        assert_eq!(parse_xlsx_cell_ref("123"), None);
        assert_eq!(parse_xlsx_cell_ref("ABC"), None);
    }

    #[test]
    fn test_extract_formulas_from_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet>
            <sheetData>
                <row r="103">
                    <c r="R103"><v>100</v></c>
                </row>
                <row r="104">
                    <c r="R104"><f>R103*1.03</f></c>
                </row>
                <row r="105">
                    <c r="R105"><f>R104*1.03</f><v>106.09</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let mut formulas = Vec::new();
        extract_formulas_from_xml(xml, 0, &mut formulas);

        // Should find both formula cells (with and without <v>)
        assert_eq!(formulas.len(), 2);

        // R104 (row=103, col=17): formula without <v>
        assert_eq!(formulas[0], (0, 103, 17, "R103*1.03".to_string()));

        // R105 (row=104, col=17): formula with <v>
        assert_eq!(formulas[1], (0, 104, 17, "R104*1.03".to_string()));
    }

    #[test]
    fn test_extract_shared_formula_followers() {
        // Simulates Excel's shared formula storage:
        // R100 is the master: <f t="shared" si="5" ref="R100:R104">R99*1.03</f>
        // R101-R104 are followers: <f t="shared" si="5"/>
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet>
            <sheetData>
                <row r="99"><c r="R99"><v>100</v></c></row>
                <row r="100"><c r="R100"><f t="shared" si="5" ref="R100:R104">R99*1.03</f><v>103</v></c></row>
                <row r="101"><c r="R101"><f t="shared" si="5"/></c></row>
                <row r="102"><c r="R102"><f t="shared" si="5"/></c></row>
                <row r="103"><c r="R103"><f t="shared" si="5"/></c></row>
                <row r="104"><c r="R104"><f t="shared" si="5"/></c></row>
            </sheetData>
        </worksheet>"#;

        let mut formulas = Vec::new();
        extract_formulas_from_xml(xml, 0, &mut formulas);

        // Should find 5 formula cells: 1 master + 4 followers
        assert_eq!(formulas.len(), 5, "Expected 5 formulas, got {:?}", formulas);

        // Master R100: R99*1.03
        assert_eq!(formulas[0], (0, 99, 17, "R99*1.03".to_string()));

        // Followers should have shifted references:
        // R101 (offset +1 row): R100*1.03
        assert_eq!(formulas[1], (0, 100, 17, "R100*1.03".to_string()));
        // R102 (offset +2 rows): R101*1.03
        assert_eq!(formulas[2], (0, 101, 17, "R101*1.03".to_string()));
        // R103 (offset +3 rows): R102*1.03
        assert_eq!(formulas[3], (0, 102, 17, "R102*1.03".to_string()));
        // R104 (offset +4 rows): R103*1.03
        assert_eq!(formulas[4], (0, 103, 17, "R103*1.03".to_string()));
    }

    #[test]
    fn test_adjust_formula_refs_for_shared() {
        // Simple row shift
        assert_eq!(
            adjust_formula_refs_for_shared("R99*1.03", 1, 0),
            "R100*1.03"
        );
        assert_eq!(
            adjust_formula_refs_for_shared("R99*1.03", 4, 0),
            "R103*1.03"
        );

        // Column shift
        assert_eq!(
            adjust_formula_refs_for_shared("A1+B1", 0, 1),
            "B1+C1"
        );

        // Both row and column
        assert_eq!(
            adjust_formula_refs_for_shared("A1*B2", 2, 1),
            "B3*C4"
        );

        // Absolute refs should not shift
        assert_eq!(
            adjust_formula_refs_for_shared("$A$1*B2", 1, 1),
            "$A$1*C3"
        );

        // Range ref
        assert_eq!(
            adjust_formula_refs_for_shared("SUM(A1:A10)", 5, 0),
            "SUM(A6:A15)"
        );

        // col_num_to_letters
        assert_eq!(col_num_to_letters(0), "A");
        assert_eq!(col_num_to_letters(25), "Z");
        assert_eq!(col_num_to_letters(26), "AA");
        assert_eq!(col_num_to_letters(27), "AB");
    }

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
            <si><t>Hello</t></si>
            <si><t>World</t></si>
            <si><r><t>Rich</t></r><r><t> Text</t></r></si>
            <si><t>10%</t></si>
        </sst>"#;
        let strings = parse_shared_strings(xml);
        assert_eq!(strings.len(), 4);
        assert_eq!(strings[0], "Hello");
        assert_eq!(strings[1], "World");
        assert_eq!(strings[2], "Rich Text");
        assert_eq!(strings[3], "10%");
    }

    #[test]
    fn test_resolve_cell_value() {
        let shared = vec!["Vacancy Rate".to_string(), "$5.00".to_string()];

        // Shared string lookup
        assert_eq!(
            resolve_cell_value(Some("s"), "0", "", &shared),
            Some("Vacancy Rate".to_string())
        );
        assert_eq!(
            resolve_cell_value(Some("s"), "1", "", &shared),
            Some("$5.00".to_string())
        );
        // Out of bounds
        assert_eq!(resolve_cell_value(Some("s"), "99", "", &shared), None);

        // Numeric
        assert_eq!(
            resolve_cell_value(Some("n"), "0.1", "", &shared),
            Some("0.1".to_string())
        );
        assert_eq!(
            resolve_cell_value(None, "42", "", &shared),
            Some("42".to_string())
        );

        // Boolean
        assert_eq!(
            resolve_cell_value(Some("b"), "1", "", &shared),
            Some("TRUE".to_string())
        );
        assert_eq!(
            resolve_cell_value(Some("b"), "0", "", &shared),
            Some("FALSE".to_string())
        );

        // Inline string
        assert_eq!(
            resolve_cell_value(Some("inlineStr"), "", "inline text", &shared),
            Some("inline text".to_string())
        );

        // Empty values
        assert_eq!(resolve_cell_value(Some("n"), "", "", &shared), None);
        assert_eq!(resolve_cell_value(Some("inlineStr"), "", "", &shared), None);
    }

    #[test]
    fn test_extract_values_from_xml() {
        let shared_strings = vec![
            "Vacancy Rate".to_string(),
            "$5.00".to_string(),
        ];
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet>
            <sheetData>
                <row r="8">
                    <c r="S8" t="s"><v>0</v></c>
                    <c r="T8" t="n"><v>0.1</v></c>
                </row>
                <row r="9">
                    <c r="S9" t="s"><v>1</v></c>
                    <c r="T9"><v>5</v></c>
                </row>
                <row r="10">
                    <c r="S10" t="inlineStr"><is><t>Inline</t></is></c>
                </row>
                <row r="11">
                    <c r="T11"><f>SUM(A1:A10)</f><v>55</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let mut out = Vec::new();
        extract_values_from_xml(xml, 0, &shared_strings, &mut out);

        // Should have 5 value cells (the formula cell T11 should be skipped)
        assert_eq!(out.len(), 5);

        // S8: shared string "Vacancy Rate"
        assert_eq!(out[0], (0, 7, 18, "Vacancy Rate".to_string()));
        // T8: numeric 0.1
        assert_eq!(out[1], (0, 7, 19, "0.1".to_string()));
        // S9: shared string "$5.00"
        assert_eq!(out[2], (0, 8, 18, "$5.00".to_string()));
        // T9: numeric 5 (no type = numeric)
        assert_eq!(out[3], (0, 8, 19, "5".to_string()));
        // S10: inline string "Inline"
        assert_eq!(out[4], (0, 9, 18, "Inline".to_string()));
        // Note: T11 has a formula, so it should NOT appear
    }

    #[test]
    fn test_xlsx_roundtrip_formatting() {
        use visigrid_engine::cell::{Alignment, CellFormat, CellBorder, BorderStyle, NumberFormat};

        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("fmt_roundtrip.xlsx");

        let mut wb = Workbook::new();
        let sheet = wb.sheet_mut(0).unwrap();

        // A1: bold + italic
        sheet.set_value(0, 0, "Bold Italic");
        let mut fmt = CellFormat::default();
        fmt.bold = true;
        fmt.italic = true;
        sheet.set_format_from_import(0, 0, fmt);

        // B1: font color (red) + font size 14
        sheet.set_value(0, 1, "Red 14pt");
        let mut fmt = CellFormat::default();
        fmt.font_color = Some([255, 0, 0, 255]); // red RGBA
        fmt.font_size = Some(14.0);
        sheet.set_format_from_import(0, 1, fmt);

        // C1: underline + strikethrough
        sheet.set_value(0, 2, "U+S");
        let mut fmt = CellFormat::default();
        fmt.underline = true;
        fmt.strikethrough = true;
        sheet.set_format_from_import(0, 2, fmt);

        // A2: CenterAcrossSelection alignment
        sheet.set_value(1, 0, "Centered Title");
        let mut fmt = CellFormat::default();
        fmt.alignment = Alignment::CenterAcrossSelection;
        sheet.set_format_from_import(1, 0, fmt);

        // B2: also CenterAcrossSelection (empty continuation cell)
        let mut fmt = CellFormat::default();
        fmt.alignment = Alignment::CenterAcrossSelection;
        sheet.set_format_from_import(1, 1, fmt);

        // A3: background color (light blue)
        sheet.set_value(2, 0, "Blue bg");
        let mut fmt = CellFormat::default();
        fmt.background_color = Some([173, 216, 230, 255]);
        sheet.set_format_from_import(2, 0, fmt);

        // B3: bottom border
        sheet.set_value(2, 1, "Border");
        let mut fmt = CellFormat::default();
        fmt.border_bottom = CellBorder { style: BorderStyle::Thin, color: Some([0, 0, 0, 255]) };
        sheet.set_format_from_import(2, 1, fmt);

        // A4: custom number format
        sheet.set_value(3, 0, "1234.5");
        let mut fmt = CellFormat::default();
        fmt.number_format = NumberFormat::Custom("#,##0.00".to_string());
        sheet.set_format_from_import(3, 0, fmt);

        // Export
        let export_result = export(&wb, &path, None).expect("Export should succeed");
        assert!(export_result.cells_exported > 0);

        // Reimport
        let (imported_wb, result) = import(&path).expect("Import should succeed");
        assert!(result.styles_imported > 0, "Should import styles");
        assert!(result.unique_styles > 0, "Should have unique styles");

        let s = &imported_wb.sheets()[0];

        // Verify bold + italic (A1)
        let f = s.get_format(0, 0);
        assert!(f.bold, "A1 should be bold");
        assert!(f.italic, "A1 should be italic");

        // Verify font color + size (B1)
        let f = s.get_format(0, 1);
        assert!(f.font_color.is_some(), "B1 should have font color");
        if let Some(color) = f.font_color {
            assert_eq!(color[0], 255, "B1 font color red channel");
            assert_eq!(color[1], 0, "B1 font color green channel");
            assert_eq!(color[2], 0, "B1 font color blue channel");
        }
        assert!(f.font_size.is_some(), "B1 should have font size");
        if let Some(size) = f.font_size {
            assert!((size - 14.0).abs() < 0.1, "B1 font size should be ~14, got {}", size);
        }

        // Verify underline + strikethrough (C1)
        let f = s.get_format(0, 2);
        assert!(f.underline, "C1 should be underlined");
        assert!(f.strikethrough, "C1 should have strikethrough");

        // Verify CenterAcrossSelection (A2)
        let f = s.get_format(1, 0);
        assert_eq!(f.alignment, Alignment::CenterAcrossSelection, "A2 should be CenterAcrossSelection");

        // Verify CenterAcrossSelection continuation (B2)
        let f = s.get_format(1, 1);
        assert_eq!(f.alignment, Alignment::CenterAcrossSelection, "B2 should be CenterAcrossSelection");

        // Verify background color (A3)
        let f = s.get_format(2, 0);
        assert!(f.background_color.is_some(), "A3 should have background color");

        // Verify border (B3)
        let f = s.get_format(2, 1);
        assert!(f.border_bottom.style != BorderStyle::None, "B3 should have bottom border");

        // Verify custom number format (A4)
        let f = s.get_format(3, 0);
        match &f.number_format {
            NumberFormat::Custom(code) => {
                assert!(code.contains("#,##0"), "A4 number format should contain '#,##0', got '{}'", code);
            }
            other => {
                // May be mapped to a built-in format, which is also acceptable
                // The key test is that the value displays correctly
                let display = s.get_formatted_display(3, 0);
                assert!(
                    display.contains("1,234") || display.contains("1234"),
                    "A4 should display formatted number, got '{}' (format: {:?})", display, other
                );
            }
        }
    }

}
