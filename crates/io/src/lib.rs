// File I/O operations

#[cfg(any(feature = "native", feature = "recipes"))]
pub mod csv;
#[cfg(any(feature = "native", feature = "recipes"))]
pub mod csv_import;
#[cfg(feature = "native")]
pub mod duckdb;
pub mod json;
pub mod content_protection;
#[cfg(feature = "native")]
pub mod native;
#[cfg(any(feature = "native", feature = "recipes"))]
pub mod parquet;
#[cfg(feature = "native")]
pub mod parquet_export;
#[cfg(any(feature = "native", feature = "recipes"))]
pub mod recipe;
#[cfg(feature = "native")]
pub mod recipe_table;
#[cfg(feature = "native")]
pub mod recipe_trust;
#[cfg(feature = "native")]
pub mod scripting;
pub mod table_recovery;
#[cfg(feature = "native")]
pub mod truth;
#[cfg(feature = "native")]
pub mod xlsx;
#[cfg(feature = "native")]
mod xlsx_comments;
#[cfg(feature = "native")]
mod xlsx_cond_formats;
#[cfg(feature = "native")]
mod xlsx_names;
#[cfg(feature = "native")]
mod xlsx_formula_cache;
#[cfg(feature = "native")]
mod xlsx_arrays;
#[cfg(feature = "native")]
mod xlsx_rich_errors;
#[cfg(feature = "native")]
mod xlsx_tables;
#[cfg(feature = "native")]
mod xlsx_table_filters;
#[cfg(feature = "native")]
mod xlsx_table_sorts;
#[cfg(feature = "native")]
mod xlsx_sorted_export;
#[cfg(feature = "native")]
pub mod xlsx_styles;
#[cfg(feature = "native")]
pub mod xlsx_validation;

/// Native .sheet format version
/// Increment when schema changes in a way that old versions can't read
pub const NATIVE_FORMAT_VERSION: u32 = 1;

/// A formula's stored result, kept aside during load.
///
/// Both file formats save a formula's computed value alongside its source. On
/// load they recompute and throw the stored value away, which is right until
/// the formula calls something this build has no definition for — a custom
/// function, which lives in the host rather than the engine. Then recomputing
/// replaces a real number with "Unknown function" and writes that back.
#[derive(Debug, Clone)]
pub enum CachedFormulaValue {
    Number(f64),
    Text(String),
}

/// Restore stored results for formulas this build could not evaluate.
///
/// Shared by the visigrid-json and .sheet loaders so the two cannot disagree
/// about when a value is kept. Narrow on purpose:
///
/// - only "Unknown function" errors. #DIV/0! and #REF! are answers this engine
///   can produce, and preserving a stale number over a real error would be
///   worse than the bug being fixed.
/// - only where the formula still matches the one the value was computed for.
/// - only where a value was actually stored; nothing is invented.
///
/// Kept cells are recorded on the sheet so the fact can be reported rather
/// than the value silently restored — it was computed elsewhere and its inputs
/// may have moved since.
pub fn keep_uncomputable_values(
    wb: &mut visigrid_engine::workbook::Workbook,
    cached: &[(usize, usize, usize, CachedFormulaValue)],
) {
    use visigrid_engine::formula::eval::Value;

    let mut kept_by_sheet: std::collections::HashMap<usize, Vec<(usize, usize)>> =
        std::collections::HashMap::new();

    for (sheet_idx, row, col, value) in cached {
        let Some(sheet) = wb.sheets().get(*sheet_idx) else { continue };
        if !sheet.get_display(*row, *col).starts_with("Unknown function") {
            continue;
        }
        match value {
            CachedFormulaValue::Number(n) => sheet.cache_computed(*row, *col, Value::Number(*n)),
            CachedFormulaValue::Text(t) => sheet.cache_computed(*row, *col, Value::Text(t.clone())),
        }
        kept_by_sheet.entry(*sheet_idx).or_default().push((*row, *col));
    }

    for (sheet_idx, cells) in kept_by_sheet {
        if let Some(sheet) = wb.sheet_mut(sheet_idx) {
            sheet.kept_uncomputable.extend(cells);
        }
    }
}
