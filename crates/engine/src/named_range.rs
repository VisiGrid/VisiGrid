//! Named range definitions and management
//!
//! Named ranges allow users to give meaningful names to cells or ranges,
//! making formulas more readable (e.g., =SUM(Revenue) instead of =SUM(A1:A100)).

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

/// A named range that maps a name to a cell reference or range
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct NamedRange {
    /// The name (case-insensitive for lookups, but preserves original case)
    pub name: String,

    /// What the name refers to
    pub target: NamedRangeTarget,

    /// Optional description for documentation
    pub description: Option<String>,
}

/// The target of a named range - either a single cell or a rectangular range
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub enum NamedRangeTarget {
    /// Single cell reference
    Cell {
        sheet: usize,
        row: usize,
        col: usize,
    },
    /// Rectangular range reference
    Range {
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    },
}

impl NamedRange {
    /// Create a new named range pointing to a single cell
    pub fn cell(name: impl Into<String>, sheet: usize, row: usize, col: usize) -> Self {
        Self {
            name: name.into(),
            target: NamedRangeTarget::Cell { sheet, row, col },
            description: None,
        }
    }

    /// Create a new named range pointing to a range
    pub fn range(
        name: impl Into<String>,
        sheet: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Self {
        Self {
            name: name.into(),
            target: NamedRangeTarget::Range {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
            },
            description: None,
        }
    }

    /// Add a description to this named range
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = Some(description.into());
        self
    }

    /// Get the cell reference string (e.g., "A1" or "A1:B10")
    pub fn reference_string(&self) -> String {
        match &self.target {
            NamedRangeTarget::Cell { row, col, .. } => {
                format!("{}{}", col_to_letter(*col), row + 1)
            }
            NamedRangeTarget::Range {
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } => {
                format!(
                    "{}{}:{}{}",
                    col_to_letter(*start_col),
                    start_row + 1,
                    col_to_letter(*end_col),
                    end_row + 1
                )
            }
        }
    }

    /// Check if this named range references the given cell
    pub fn references_cell(&self, sheet: usize, row: usize, col: usize) -> bool {
        match &self.target {
            NamedRangeTarget::Cell {
                sheet: s,
                row: r,
                col: c,
            } => *s == sheet && *r == row && *c == col,
            NamedRangeTarget::Range {
                sheet: s,
                start_row,
                start_col,
                end_row,
                end_col,
            } => {
                *s == sheet
                    && row >= *start_row
                    && row <= *end_row
                    && col >= *start_col
                    && col <= *end_col
            }
        }
    }
}

/// Validate a named range identifier
/// Rules:
/// - Must start with letter or underscore
/// - Can contain letters, numbers, underscores, and dots (for namespaces)
/// - Cannot be a cell reference (A1, BC23)
/// - Cannot be a range (A1:B2)
/// - Cannot be a function name (SUM, IF, VLOOKUP)
/// - Cannot be a boolean or error literal (TRUE, FALSE, #REF!)
pub fn is_valid_name(name: &str) -> Result<(), String> {
    // Trim whitespace
    let name = name.trim();

    if name.is_empty() {
        return Err("Name cannot be empty".into());
    }

    let first = name.chars().next().unwrap();

    // Must start with letter or underscore (not digit)
    if first.is_ascii_digit() {
        return Err("Name must start with a letter or underscore, not a digit".into());
    }

    if !first.is_alphabetic() && first != '_' {
        return Err("Name must start with a letter or underscore".into());
    }

    // Valid characters: letters, numbers, underscores, dots (for namespaces like ACME.Revenue)
    if !name.chars().all(|c| c.is_alphanumeric() || c == '_' || c == '.') {
        return Err("Name can only contain letters, numbers, underscores, and dots".into());
    }

    // Cannot end with a dot
    if name.ends_with('.') {
        return Err("Name cannot end with a dot".into());
    }

    // Cannot have consecutive dots
    if name.contains("..") {
        return Err("Name cannot have consecutive dots".into());
    }

    let upper = name.to_uppercase();

    // Check if it looks like a cell reference (e.g., A1, AB123)
    if looks_like_cell_ref(name) {
        return Err(format!(
            "'{}' looks like a cell reference (e.g., A1, BC23). Choose a different name.",
            name
        ));
    }

    // Check if it looks like a range (e.g., A1:B2)
    if looks_like_range(name) {
        return Err(format!(
            "'{}' looks like a range reference (e.g., A1:B2). Choose a different name.",
            name
        ));
    }

    // Check for boolean literals
    if upper == "TRUE" || upper == "FALSE" {
        return Err(format!(
            "'{}' is a reserved boolean value. Choose a different name.",
            name
        ));
    }

    // Check for error literals (with or without #)
    let error_literals = [
        "REF", "DIV", "NAME", "VALUE", "NUM", "NA", "NULL", "ERROR",
        "#REF!", "#DIV/0!", "#NAME?", "#VALUE!", "#NUM!", "#N/A", "#NULL!", "#ERROR!",
    ];
    if error_literals.iter().any(|e| upper == *e || upper == e.replace(['#', '!', '?', '/'], "")) {
        return Err(format!(
            "'{}' conflicts with an error value. Choose a different name.",
            name
        ));
    }

    // Check for function names (comprehensive list of common spreadsheet functions)
    if is_function_name(&upper) {
        return Err(format!(
            "'{}' is a function name. Choose a different name to avoid confusion.",
            name
        ));
    }

    Ok(())
}

/// Check if name matches a known spreadsheet function (case-insensitive)
fn is_function_name(upper_name: &str) -> bool {
    // Comprehensive list of Excel/spreadsheet function names
    // Organized by category for maintainability
    const FUNCTIONS: &[&str] = &[
        // Math & Trig (30+)
        "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ",
        "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS",
        "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS",
        "MIN", "MINA", "MAX", "MAXA", "MEDIAN", "MODE",
        "ABS", "ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC", "INT", "FLOOR", "CEILING",
        "MOD", "POWER", "SQRT", "EXP", "LN", "LOG", "LOG10",
        "PI", "RAND", "RANDBETWEEN",
        "SIN", "COS", "TAN", "ASIN", "ACOS", "ATAN", "ATAN2",
        "DEGREES", "RADIANS", "SIGN",
        "PRODUCT", "QUOTIENT", "GCD", "LCM", "FACT", "COMBIN", "PERMUT",

        // Text (25+)
        "LEFT", "RIGHT", "MID", "LEN", "FIND", "SEARCH",
        "CONCAT", "CONCATENATE", "TEXTJOIN",
        "UPPER", "LOWER", "PROPER", "TRIM", "CLEAN",
        "SUBSTITUTE", "REPLACE", "REPT",
        "TEXT", "VALUE", "FIXED", "DOLLAR",
        "CHAR", "CODE", "UNICODE", "UNICHAR",
        "EXACT", "T", "N",

        // Logical (10+)
        "IF", "IFS", "IFERROR", "IFNA",
        "AND", "OR", "NOT", "XOR",
        "SWITCH", "CHOOSE",

        // Lookup & Reference (20+)
        "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP",
        "INDEX", "MATCH", "XMATCH",
        "INDIRECT", "OFFSET", "ADDRESS",
        "ROW", "ROWS", "COLUMN", "COLUMNS",
        "TRANSPOSE", "SORT", "SORTBY", "UNIQUE", "FILTER",
        "HYPERLINK",

        // Date & Time (25+)
        "DATE", "DATEVALUE", "TIME", "TIMEVALUE",
        "NOW", "TODAY", "YEAR", "MONTH", "DAY",
        "HOUR", "MINUTE", "SECOND",
        "WEEKDAY", "WEEKNUM", "ISOWEEKNUM",
        "DAYS", "DAYS360", "NETWORKDAYS", "WORKDAY",
        "EDATE", "EOMONTH", "DATEDIF",
        "YEARFRAC",

        // Statistical (20+)
        "STDEV", "STDEVA", "STDEVP", "STDEVPA",
        "VAR", "VARA", "VARP", "VARPA",
        "LARGE", "SMALL", "RANK", "PERCENTILE", "QUARTILE",
        "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
        "CORREL", "COVAR", "SLOPE", "INTERCEPT", "FORECAST",
        "TREND", "GROWTH", "LINEST",

        // Financial (15+)
        "NPV", "IRR", "XNPV", "XIRR",
        "PMT", "PPMT", "IPMT", "FV", "PV", "NPER", "RATE",
        "SLN", "DB", "DDB", "SYD",

        // Information (15+)
        "ISBLANK", "ISERROR", "ISERR", "ISNA", "ISTEXT", "ISNUMBER",
        "ISLOGICAL", "ISREF", "ISFORMULA", "ISEVEN", "ISODD",
        "TYPE", "CELL", "INFO", "SHEET", "SHEETS",
        "ERROR.TYPE",

        // Engineering & Conversion
        "CONVERT", "DEC2BIN", "DEC2HEX", "DEC2OCT",
        "BIN2DEC", "BIN2HEX", "BIN2OCT",
        "HEX2DEC", "HEX2BIN", "HEX2OCT",
        "OCT2DEC", "OCT2BIN", "OCT2HEX",

        // Array/Dynamic
        "SEQUENCE", "RANDARRAY", "LET", "LAMBDA",
        "MAP", "REDUCE", "SCAN", "MAKEARRAY", "BYROW", "BYCOL",
        "HSTACK", "VSTACK", "TOROW", "TOCOL", "WRAPROWS", "WRAPCOLS",
        "TAKE", "DROP", "EXPAND", "CHOOSEROWS", "CHOOSECOLS",

        // Database
        "DSUM", "DAVERAGE", "DCOUNT", "DCOUNTA", "DMAX", "DMIN",
        "DGET", "DPRODUCT", "DSTDEV", "DSTDEVP", "DVAR", "DVARP",

        // Other common
        "AGGREGATE", "SUBTOTAL", "GETPIVOTDATA",
    ];

    FUNCTIONS.contains(&upper_name)
}

/// Check if a string looks like a range reference (e.g., A1:B2, $A$1:$B$2)
fn looks_like_range(s: &str) -> bool {
    // Remove any $ signs for absolute reference checking
    let clean = s.replace('$', "");

    // Must contain exactly one colon
    let parts: Vec<&str> = clean.split(':').collect();
    if parts.len() != 2 {
        return false;
    }

    // Both parts must look like cell references
    looks_like_cell_ref(parts[0]) && looks_like_cell_ref(parts[1])
}

/// Check if a string looks like a cell reference (e.g., A1, AB123, XFD1048576)
/// Only rejects names that could be interpreted as actual cell references:
/// - Column part: 1-3 ASCII letters that form a valid column (A-XFD)
/// - Row part: digits forming a valid row number
fn looks_like_cell_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();

    // Collect the letter part (column)
    let mut col_str = String::new();
    while chars.peek().map(|c| c.is_ascii_alphabetic()).unwrap_or(false) {
        col_str.push(chars.next().unwrap());
    }

    // Must have 1-3 letters to be a valid column
    if col_str.is_empty() || col_str.len() > 3 {
        return false;
    }

    // Check if the column is within valid range (A=1 to XFD=16384)
    let col_num = col_str.to_uppercase().chars().fold(0u32, |acc, c| {
        acc * 26 + (c as u32 - 'A' as u32 + 1)
    });
    if col_num > 16384 {
        return false;
    }

    // Collect the digit part (row)
    let mut row_str = String::new();
    while chars.peek().map(|c| c.is_ascii_digit()).unwrap_or(false) {
        row_str.push(chars.next().unwrap());
    }

    // Must have digits and nothing else remaining
    if row_str.is_empty() || chars.next().is_some() {
        return false;
    }

    // Check if row is within valid range (1 to 1048576 for Excel, we'll use reasonable limit)
    if let Ok(row_num) = row_str.parse::<u32>() {
        row_num >= 1 && row_num <= 1048576
    } else {
        false // Number too large to parse
    }
}

/// Convert column index to letter(s) (0 = A, 25 = Z, 26 = AA, etc.)
fn col_to_letter(col: usize) -> String {
    let mut s = String::new();
    let mut n = col;
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    s
}

/// Storage for named ranges in a workbook
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NamedRangeStore {
    /// Named ranges keyed by lowercase name for case-insensitive lookup
    ranges: HashMap<String, NamedRange>,
}

impl NamedRangeStore {
    pub fn new() -> Self {
        Self::default()
    }

    /// Add or update a named range
    pub fn set(&mut self, range: NamedRange) -> Result<(), String> {
        is_valid_name(&range.name)?;
        self.ranges.insert(range.name.to_lowercase(), range);
        Ok(())
    }

    /// Get a named range by name (case-insensitive)
    pub fn get(&self, name: &str) -> Option<&NamedRange> {
        self.ranges.get(&name.to_lowercase())
    }

    /// Remove a named range by name (case-insensitive)
    pub fn remove(&mut self, name: &str) -> Option<NamedRange> {
        self.ranges.remove(&name.to_lowercase())
    }

    /// Update the description of a named range
    pub fn set_description(&mut self, name: &str, description: Option<String>) -> Result<(), String> {
        let key = name.to_lowercase();
        if let Some(range) = self.ranges.get_mut(&key) {
            range.description = description;
            Ok(())
        } else {
            Err(format!("Name '{}' not found", name))
        }
    }

    /// Rename a named range (returns error if old name doesn't exist or new name is invalid/taken)
    pub fn rename(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        is_valid_name(new_name)?;

        let old_key = old_name.to_lowercase();
        let new_key = new_name.to_lowercase();

        if old_key != new_key && self.ranges.contains_key(&new_key) {
            return Err(format!("Name '{}' already exists", new_name));
        }

        if let Some(mut range) = self.ranges.remove(&old_key) {
            range.name = new_name.to_string();
            self.ranges.insert(new_key, range);
            Ok(())
        } else {
            Err(format!("Name '{}' not found", old_name))
        }
    }

    /// Check if a name exists (case-insensitive)
    pub fn contains(&self, name: &str) -> bool {
        self.ranges.contains_key(&name.to_lowercase())
    }

    /// List all named ranges
    pub fn list(&self) -> Vec<&NamedRange> {
        self.ranges.values().collect()
    }

    /// Find all named ranges that reference a specific cell
    pub fn find_by_cell(&self, sheet: usize, row: usize, col: usize) -> Vec<&NamedRange> {
        self.ranges
            .values()
            .filter(|nr| nr.references_cell(sheet, row, col))
            .collect()
    }

    /// Get the number of named ranges
    pub fn len(&self) -> usize {
        self.ranges.len()
    }

    /// Check if there are no named ranges
    pub fn is_empty(&self) -> bool {
        self.ranges.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_valid_names() {
        assert!(is_valid_name("Revenue").is_ok());
        assert!(is_valid_name("_private").is_ok());
        assert!(is_valid_name("Sales2024").is_ok());
        assert!(is_valid_name("total_cost").is_ok());
    }

    #[test]
    fn test_invalid_names_basic() {
        assert!(is_valid_name("").is_err());
        assert!(is_valid_name("   ").is_err()); // whitespace only
        assert!(is_valid_name("123abc").is_err()); // starts with digit
        assert!(is_valid_name("has space").is_err());
        assert!(is_valid_name("has-dash").is_err());
        assert!(is_valid_name("has@symbol").is_err());
    }

    #[test]
    fn test_valid_names_with_dots() {
        // Dots allowed for namespacing
        assert!(is_valid_name("ACME.Revenue").is_ok());
        assert!(is_valid_name("Sales.Q1.Total").is_ok());
        // But not invalid dot usage
        assert!(is_valid_name(".StartDot").is_err());
        assert!(is_valid_name("EndDot.").is_err());
        assert!(is_valid_name("Double..Dot").is_err());
    }

    #[test]
    fn test_cell_reference_blocking() {
        // Common cell references should be blocked
        assert!(is_valid_name("A1").is_err());
        assert!(is_valid_name("a1").is_err()); // case insensitive
        assert!(is_valid_name("AB123").is_err());
        assert!(is_valid_name("XFD1048576").is_err()); // max Excel cell
        // But valid names that could look like cell refs
        assert!(is_valid_name("XGA1").is_ok()); // beyond valid column range
        assert!(is_valid_name("AAAA1").is_ok()); // 4 letters = ok
        assert!(is_valid_name("Revenue1").is_ok()); // not a valid cell pattern
    }

    #[test]
    fn test_range_reference_blocking() {
        // Ranges should be blocked
        assert!(is_valid_name("A1:B2").is_err());
        assert!(is_valid_name("$A$1:$B$2").is_err()); // absolute references
        assert!(is_valid_name("A1:B2").is_err());
        // But colons in other contexts should fail on character rules anyway
        assert!(is_valid_name("Data:2024").is_err()); // colon not allowed
    }

    #[test]
    fn test_function_name_blocking() {
        // Common functions should be blocked
        assert!(is_valid_name("SUM").is_err());
        assert!(is_valid_name("sum").is_err()); // case insensitive
        assert!(is_valid_name("VLOOKUP").is_err());
        assert!(is_valid_name("IF").is_err());
        assert!(is_valid_name("Average").is_err());
        assert!(is_valid_name("COUNT").is_err());
        assert!(is_valid_name("INDEX").is_err());
        assert!(is_valid_name("MATCH").is_err());
        // Check error message is helpful
        let err = is_valid_name("SUM").unwrap_err();
        assert!(err.contains("function name"), "Error should mention function: {}", err);
    }

    #[test]
    fn test_boolean_literal_blocking() {
        assert!(is_valid_name("TRUE").is_err());
        assert!(is_valid_name("true").is_err());
        assert!(is_valid_name("FALSE").is_err());
        assert!(is_valid_name("false").is_err());
        // Check error message
        let err = is_valid_name("TRUE").unwrap_err();
        assert!(err.contains("boolean"), "Error should mention boolean: {}", err);
    }

    #[test]
    fn test_error_literal_blocking() {
        // Error values should be blocked
        assert!(is_valid_name("REF").is_err());
        assert!(is_valid_name("NA").is_err());
        assert!(is_valid_name("VALUE").is_err());
        assert!(is_valid_name("DIV").is_err());
        assert!(is_valid_name("NUM").is_err());
        // Note: #REF! etc. would fail on # character anyway
    }

    #[test]
    fn test_whitespace_trimming() {
        // Leading/trailing whitespace should be trimmed
        assert!(is_valid_name("  Revenue  ").is_ok());
        assert!(is_valid_name("\tMyName\n").is_ok());
    }

    #[test]
    fn test_named_range_store() {
        let mut store = NamedRangeStore::new();

        // Add a named range
        store.set(NamedRange::cell("Revenue", 0, 0, 0)).unwrap();
        assert!(store.contains("Revenue"));
        assert!(store.contains("revenue")); // Case insensitive

        // Get it back
        let nr = store.get("REVENUE").unwrap();
        assert_eq!(nr.name, "Revenue"); // Preserves original case

        // Rename it
        store.rename("Revenue", "TotalRevenue").unwrap();
        assert!(!store.contains("Revenue"));
        assert!(store.contains("TotalRevenue"));

        // Remove it
        store.remove("totalrevenue");
        assert!(store.is_empty());
    }

    #[test]
    fn test_reference_string() {
        let cell = NamedRange::cell("Test", 0, 0, 0);
        assert_eq!(cell.reference_string(), "A1");

        let cell = NamedRange::cell("Test", 0, 99, 27);
        assert_eq!(cell.reference_string(), "AB100");

        let range = NamedRange::range("Test", 0, 0, 0, 9, 2);
        assert_eq!(range.reference_string(), "A1:C10");
    }

    #[test]
    fn test_references_cell() {
        let cell = NamedRange::cell("Test", 0, 5, 3);
        assert!(cell.references_cell(0, 5, 3));
        assert!(!cell.references_cell(0, 5, 4));
        assert!(!cell.references_cell(1, 5, 3)); // Different sheet

        let range = NamedRange::range("Test", 0, 0, 0, 9, 2);
        assert!(range.references_cell(0, 0, 0));
        assert!(range.references_cell(0, 5, 1));
        assert!(range.references_cell(0, 9, 2));
        assert!(!range.references_cell(0, 10, 0)); // Outside range
        assert!(!range.references_cell(0, 0, 3)); // Outside range
    }

    #[test]
    fn test_rename_case_only() {
        // Case-only rename should succeed (foo -> Foo)
        let mut store = NamedRangeStore::new();
        store.set(NamedRange::cell("foo", 0, 0, 0)).unwrap();

        // Rename to different case
        let result = store.rename("foo", "Foo");
        assert!(result.is_ok(), "Case-only rename should succeed");

        // Should still have exactly one entry
        assert_eq!(store.len(), 1);

        // The name should be updated to the new case
        let nr = store.get("foo").unwrap();
        assert_eq!(nr.name, "Foo");

        // Also works with FOO -> fOo -> FoO
        store.rename("Foo", "fOo").unwrap();
        assert_eq!(store.get("foo").unwrap().name, "fOo");
        store.rename("foo", "FoO").unwrap();
        assert_eq!(store.get("foo").unwrap().name, "FoO");
    }

    #[test]
    fn test_rename_to_existing() {
        // Renaming to an existing name should fail
        let mut store = NamedRangeStore::new();
        store.set(NamedRange::cell("Alpha", 0, 0, 0)).unwrap();
        store.set(NamedRange::cell("Beta", 0, 1, 0)).unwrap();

        // Try to rename "Alpha" to "Beta" - should fail
        let result = store.rename("Alpha", "Beta");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(err.contains("already exists"), "Error should mention existing name: {}", err);

        // Also case-insensitive: "Alpha" -> "beta"
        let result = store.rename("Alpha", "beta");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(err.contains("already exists"), "Error should be case-insensitive: {}", err);

        // Original entries should be unchanged
        assert!(store.contains("Alpha"));
        assert!(store.contains("Beta"));
        assert_eq!(store.len(), 2);
    }

    #[test]
    fn test_rename_nonexistent() {
        // Renaming a non-existent name should fail
        let mut store = NamedRangeStore::new();
        store.set(NamedRange::cell("Existing", 0, 0, 0)).unwrap();

        let result = store.rename("NonExistent", "NewName");
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(err.contains("not found"), "Error should mention name not found: {}", err);
    }

    #[test]
    fn test_rename_to_invalid_name() {
        // Renaming to an invalid name should fail
        let mut store = NamedRangeStore::new();
        store.set(NamedRange::cell("Valid", 0, 0, 0)).unwrap();

        // Try to rename to a function name
        let result = store.rename("Valid", "SUM");
        assert!(result.is_err());

        // Try to rename to a cell reference
        let result = store.rename("Valid", "A1");
        assert!(result.is_err());

        // Try to rename to an invalid character
        let result = store.rename("Valid", "has space");
        assert!(result.is_err());

        // Original should be unchanged
        assert!(store.contains("Valid"));
        assert_eq!(store.get("Valid").unwrap().name, "Valid");
    }

    #[test]
    fn test_set_description() {
        let mut store = NamedRangeStore::new();
        store.set(NamedRange::cell("Test", 0, 0, 0)).unwrap();

        // Initially no description
        assert!(store.get("Test").unwrap().description.is_none());

        // Add description
        store.set_description("Test", Some("A test range".to_string())).unwrap();
        assert_eq!(
            store.get("Test").unwrap().description,
            Some("A test range".to_string())
        );

        // Clear description
        store.set_description("Test", None).unwrap();
        assert!(store.get("Test").unwrap().description.is_none());

        // Set on non-existent should fail
        let result = store.set_description("NonExistent", Some("desc".to_string()));
        assert!(result.is_err());
    }
}
