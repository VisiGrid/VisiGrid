// AI context extraction
//
// Extracts cell data from selected range for AI prompts.
// Respects privacy mode and caps data to prevent excessive token usage.

use visigrid_engine::sheet::Sheet;

/// Maximum rows to include in AI context
const MAX_CONTEXT_ROWS: usize = 200;

/// Maximum columns to include in AI context
const MAX_CONTEXT_COLS: usize = 20;

/// Warning about context extraction
#[derive(Debug, Clone)]
pub enum ContextWarning {
    /// Selection was truncated due to size limits
    Truncated {
        original_rows: usize,
        original_cols: usize,
        actual_rows: usize,
        actual_cols: usize,
    },
    /// Selection is empty
    EmptySelection,
}

impl ContextWarning {
    pub fn message(&self) -> String {
        match self {
            ContextWarning::Truncated { original_rows, original_cols, actual_rows, actual_cols } => {
                format!(
                    "Selection truncated from {}x{} to {}x{} (max {} rows, {} cols)",
                    original_rows, original_cols, actual_rows, actual_cols,
                    MAX_CONTEXT_ROWS, MAX_CONTEXT_COLS
                )
            }
            ContextWarning::EmptySelection => "Selection is empty".to_string(),
        }
    }
}

/// Context extracted from spreadsheet for AI prompt
#[derive(Debug, Clone)]
pub struct AIContext {
    /// Sheet name
    pub sheet_name: String,

    /// Original selection range (A1 notation)
    pub original_range: String,

    /// Actual range used (may be truncated)
    pub actual_range: String,

    /// Summary for display (e.g., "Using Sheet1!A1:F200 (truncated from A1:F5000)")
    pub summary: String,

    /// Column headers (if first row looks like headers)
    pub headers: Option<Vec<String>>,

    /// Cell data as rows of values
    /// Each row is a vector of cell display values
    pub data: Vec<Vec<String>>,

    /// Number of rows in data
    pub row_count: usize,

    /// Number of columns in data
    pub col_count: usize,

    /// Warnings about extraction
    pub warnings: Vec<ContextWarning>,

    /// Whether privacy mode was active
    pub privacy_mode: bool,

    /// Top-left cell of actual range (0-indexed)
    pub top_left: (usize, usize),
}

impl AIContext {
    /// Format context as structured text for AI prompt
    pub fn to_prompt_text(&self) -> String {
        let mut result = String::new();

        // Header info
        result.push_str(&format!("Sheet: {}\n", self.sheet_name));
        result.push_str(&format!("Range: {}\n", self.actual_range));
        if self.actual_range != self.original_range {
            result.push_str(&format!("(truncated from {})\n", self.original_range));
        }
        result.push_str(&format!("Size: {} rows x {} columns\n\n", self.row_count, self.col_count));

        // Data as TSV-like format
        if let Some(headers) = &self.headers {
            result.push_str("Headers: ");
            result.push_str(&headers.join("\t"));
            result.push('\n');
        }

        result.push_str("Data:\n");
        for row in &self.data {
            result.push_str(&row.join("\t"));
            result.push('\n');
        }

        result
    }
}

/// Convert 0-indexed column to letter (0 = A, 25 = Z, 26 = AA, etc.)
pub fn col_to_letter(col: usize) -> String {
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

/// Format cell reference in A1 notation
pub fn cell_ref(row: usize, col: usize) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

/// Format range in A1 notation
pub fn range_ref(start_row: usize, start_col: usize, end_row: usize, end_col: usize) -> String {
    if start_row == end_row && start_col == end_col {
        cell_ref(start_row, start_col)
    } else {
        format!("{}:{}", cell_ref(start_row, start_col), cell_ref(end_row, end_col))
    }
}

/// Find the current region (contiguous non-empty block) around a cell.
/// Returns (start_row, start_col, end_row, end_col) 0-indexed, inclusive.
pub fn find_current_region(sheet: &Sheet, row: usize, col: usize) -> (usize, usize, usize, usize) {
    // Start from the given cell and expand outward until we hit empty rows/cols
    let mut min_row = row;
    let mut max_row = row;
    let mut min_col = col;
    let mut max_col = col;

    // Expand upward
    while min_row > 0 {
        if is_row_empty_in_range(sheet, min_row - 1, min_col, max_col) {
            break;
        }
        min_row -= 1;
        // Re-expand columns for new row
        while min_col > 0 && !sheet.get_display(min_row, min_col - 1).is_empty() {
            min_col -= 1;
        }
        while !sheet.get_display(min_row, max_col + 1).is_empty() {
            max_col += 1;
        }
    }

    // Expand downward
    loop {
        if is_row_empty_in_range(sheet, max_row + 1, min_col, max_col) {
            break;
        }
        max_row += 1;
        // Re-expand columns for new row
        while min_col > 0 && !sheet.get_display(max_row, min_col - 1).is_empty() {
            min_col -= 1;
        }
        while !sheet.get_display(max_row, max_col + 1).is_empty() {
            max_col += 1;
        }
    }

    // Expand left
    while min_col > 0 {
        if is_col_empty_in_range(sheet, min_col - 1, min_row, max_row) {
            break;
        }
        min_col -= 1;
    }

    // Expand right
    loop {
        if is_col_empty_in_range(sheet, max_col + 1, min_row, max_row) {
            break;
        }
        max_col += 1;
    }

    (min_row, min_col, max_row, max_col)
}

fn is_row_empty_in_range(sheet: &Sheet, row: usize, start_col: usize, end_col: usize) -> bool {
    for col in start_col..=end_col {
        if !sheet.get_display(row, col).is_empty() {
            return false;
        }
    }
    true
}

fn is_col_empty_in_range(sheet: &Sheet, col: usize, start_row: usize, end_row: usize) -> bool {
    for row in start_row..=end_row {
        if !sheet.get_display(row, col).is_empty() {
            return false;
        }
    }
    true
}

/// Find the used range of the sheet (bounding box of all non-empty cells).
/// Returns (start_row, start_col, end_row, end_col) 0-indexed, inclusive.
/// Returns (0, 0, 0, 0) if sheet is empty.
pub fn find_used_range(sheet: &Sheet) -> (usize, usize, usize, usize) {
    let mut min_row = usize::MAX;
    let mut max_row = 0usize;
    let mut min_col = usize::MAX;
    let mut max_col = 0usize;
    let mut found_any = false;

    // Scan within reasonable bounds (don't scan entire 65536x256)
    // Use sheet's internal knowledge if available, otherwise scan first 1000x100
    let scan_rows = 1000;
    let scan_cols = 100;

    for row in 0..scan_rows {
        for col in 0..scan_cols {
            if !sheet.get_display(row, col).is_empty() {
                found_any = true;
                min_row = min_row.min(row);
                max_row = max_row.max(row);
                min_col = min_col.min(col);
                max_col = max_col.max(col);
            }
        }
    }

    if found_any {
        (min_row, min_col, max_row, max_col)
    } else {
        (0, 0, 0, 0)
    }
}

/// Check if a row looks like headers (contains mostly text, no numbers)
fn looks_like_header_row(row: &[String]) -> bool {
    if row.is_empty() {
        return false;
    }

    let non_empty: Vec<_> = row.iter().filter(|s| !s.is_empty()).collect();
    if non_empty.is_empty() {
        return false;
    }

    // Header row: mostly non-numeric text
    let text_count = non_empty.iter()
        .filter(|s| s.parse::<f64>().is_err())
        .count();

    // At least 60% should be text (allows mixed headers like "ID", "123", "Product")
    text_count * 10 >= non_empty.len() * 6
}

/// Build AI context from a selected range
///
/// # Arguments
/// * `sheet` - The sheet to extract from
/// * `sheet_name` - Name of the sheet
/// * `start_row`, `start_col` - Top-left of selection (0-indexed)
/// * `end_row`, `end_col` - Bottom-right of selection (0-indexed, inclusive)
/// * `privacy_mode` - If true, only include cell values (no formulas)
pub fn build_ai_context(
    sheet: &Sheet,
    sheet_name: &str,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    privacy_mode: bool,
) -> AIContext {
    let mut warnings = Vec::new();

    // Normalize range (ensure start <= end)
    let (min_row, max_row) = if start_row <= end_row { (start_row, end_row) } else { (end_row, start_row) };
    let (min_col, max_col) = if start_col <= end_col { (start_col, end_col) } else { (end_col, start_col) };

    let original_rows = max_row - min_row + 1;
    let original_cols = max_col - min_col + 1;

    // Check for empty selection
    if original_rows == 0 || original_cols == 0 {
        warnings.push(ContextWarning::EmptySelection);
        return AIContext {
            sheet_name: sheet_name.to_string(),
            original_range: range_ref(min_row, min_col, max_row, max_col),
            actual_range: range_ref(min_row, min_col, max_row, max_col),
            summary: format!("Empty selection in {}", sheet_name),
            headers: None,
            data: Vec::new(),
            row_count: 0,
            col_count: 0,
            warnings,
            privacy_mode,
            top_left: (min_row, min_col),
        };
    }

    // Apply caps
    let actual_rows = original_rows.min(MAX_CONTEXT_ROWS);
    let actual_cols = original_cols.min(MAX_CONTEXT_COLS);
    let actual_max_row = min_row + actual_rows - 1;
    let actual_max_col = min_col + actual_cols - 1;

    // Record truncation warning
    if actual_rows < original_rows || actual_cols < original_cols {
        warnings.push(ContextWarning::Truncated {
            original_rows,
            original_cols,
            actual_rows,
            actual_cols,
        });
    }

    // Extract data
    let mut data = Vec::with_capacity(actual_rows);
    for row in min_row..=actual_max_row {
        let mut row_data = Vec::with_capacity(actual_cols);
        for col in min_col..=actual_max_col {
            // Always use display value (computed result)
            // Privacy mode just means we don't include formulas elsewhere
            let value = sheet.get_display(row, col);
            row_data.push(value);
        }
        data.push(row_data);
    }

    // Check for headers
    let headers = if !data.is_empty() && looks_like_header_row(&data[0]) {
        Some(data[0].clone())
    } else {
        None
    };

    // Build ranges
    let original_range = range_ref(min_row, min_col, max_row, max_col);
    let actual_range = range_ref(min_row, min_col, actual_max_row, actual_max_col);

    // Build summary
    let summary = if actual_range != original_range {
        format!(
            "Using {}!{} (truncated from {})",
            sheet_name, actual_range, original_range
        )
    } else {
        format!("Using {}!{}", sheet_name, actual_range)
    };

    AIContext {
        sheet_name: sheet_name.to_string(),
        original_range,
        actual_range,
        summary,
        headers,
        data,
        row_count: actual_rows,
        col_count: actual_cols,
        warnings,
        privacy_mode,
        top_left: (min_row, min_col),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(701), "ZZ");
        assert_eq!(col_to_letter(702), "AAA");
    }

    #[test]
    fn test_cell_ref() {
        assert_eq!(cell_ref(0, 0), "A1");
        assert_eq!(cell_ref(9, 2), "C10");
        assert_eq!(cell_ref(0, 26), "AA1");
    }

    #[test]
    fn test_range_ref() {
        assert_eq!(range_ref(0, 0, 0, 0), "A1");
        assert_eq!(range_ref(0, 0, 9, 4), "A1:E10");
    }

    #[test]
    fn test_looks_like_header_row() {
        assert!(looks_like_header_row(&["Name".to_string(), "Date".to_string(), "Amount".to_string()]));
        assert!(!looks_like_header_row(&["100".to_string(), "200".to_string(), "300".to_string()]));
        assert!(looks_like_header_row(&["ID".to_string(), "123".to_string(), "Product".to_string()])); // 66% text, borderline
        assert!(!looks_like_header_row(&[]));
    }
}
