//! Role-based auto-styling for agent-built spreadsheets
//!
//! Agents declare intent via `meta("A3:I3", { role = "header" })`.
//! GUI renders presentation (colors, borders, formats) based on role.
//! This keeps Lua scripts clean and agent output consistent.

use std::collections::HashMap;
use gpui::{Hsla, hsla};

/// Supported roles for auto-styling
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Role {
    Title,
    Header,
    Label,
    Input,
    Currency,
    Percent,
    Integer,
    Total,
    CheckResult,
}

impl Role {
    /// Parse role from string (case-insensitive)
    pub fn from_str(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "title" => Some(Role::Title),
            "header" => Some(Role::Header),
            "label" => Some(Role::Label),
            "input" => Some(Role::Input),
            "currency" => Some(Role::Currency),
            "percent" => Some(Role::Percent),
            "integer" => Some(Role::Integer),
            "total" => Some(Role::Total),
            "check_result" | "check" => Some(Role::CheckResult),
            _ => None,
        }
    }
}

/// Computed style for a role
#[derive(Debug, Clone, Default)]
pub struct RoleStyle {
    pub background: Option<Hsla>,
    pub text_color: Option<Hsla>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub align_right: bool,
    pub align_center: bool,
    pub border_bottom: bool,
    pub border_top: bool,
    /// Number format hint for display
    pub number_format: Option<NumberDisplayFormat>,
}

/// Display format for numbers (presentation only, doesn't change stored value)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberDisplayFormat {
    Currency,    // $1,234.56
    Percent,     // 3.50%
    Integer,     // 1,234
}

/// Maps roles to presentation styles
pub struct RoleStyleMap {
    styles: HashMap<Role, RoleStyle>,
}

impl Default for RoleStyleMap {
    fn default() -> Self {
        Self::new()
    }
}

impl RoleStyleMap {
    pub fn new() -> Self {
        let mut styles = HashMap::new();

        // Title: large bold
        styles.insert(Role::Title, RoleStyle {
            bold: Some(true),
            ..Default::default()
        });

        // Header: bold, centered, bottom border, light gray background
        styles.insert(Role::Header, RoleStyle {
            background: Some(hsla(0.0, 0.0, 0.92, 1.0)), // light gray
            bold: Some(true),
            align_center: true,
            border_bottom: true,
            ..Default::default()
        });

        // Label: left-aligned (default), slightly muted
        styles.insert(Role::Label, RoleStyle {
            text_color: Some(hsla(0.0, 0.0, 0.4, 1.0)), // muted gray
            ..Default::default()
        });

        // Input: blue background, white text
        styles.insert(Role::Input, RoleStyle {
            background: Some(hsla(0.6, 0.7, 0.45, 1.0)), // blue
            text_color: Some(hsla(0.0, 0.0, 1.0, 1.0)),  // white
            bold: Some(true),
            ..Default::default()
        });

        // Currency: right-aligned, currency format
        styles.insert(Role::Currency, RoleStyle {
            align_right: true,
            number_format: Some(NumberDisplayFormat::Currency),
            ..Default::default()
        });

        // Percent: right-aligned, percent format
        styles.insert(Role::Percent, RoleStyle {
            align_right: true,
            number_format: Some(NumberDisplayFormat::Percent),
            ..Default::default()
        });

        // Integer: right-aligned, no decimals
        styles.insert(Role::Integer, RoleStyle {
            align_right: true,
            number_format: Some(NumberDisplayFormat::Integer),
            ..Default::default()
        });

        // Total: bold, top border
        styles.insert(Role::Total, RoleStyle {
            bold: Some(true),
            border_top: true,
            ..Default::default()
        });

        // CheckResult: special handling in render (PASS=green, FAIL=red)
        styles.insert(Role::CheckResult, RoleStyle {
            bold: Some(true),
            ..Default::default()
        });

        Self { styles }
    }

    /// Get style for a role
    pub fn get(&self, role: Role) -> Option<&RoleStyle> {
        self.styles.get(&role)
    }

    /// Get style by role name string
    pub fn get_by_name(&self, name: &str) -> Option<&RoleStyle> {
        Role::from_str(name).and_then(|r| self.styles.get(&r))
    }
}

/// Cell metadata storage for GUI
/// Maps cell address (e.g., "A3") to key-value pairs
pub type CellMetadataMap = HashMap<String, HashMap<String, String>>;

/// Resolve metadata for a specific cell position
/// Handles both exact cell matches and range matches
pub fn get_cell_role(metadata: &CellMetadataMap, row: usize, col: usize) -> Option<Role> {
    let cell_addr = format_cell_address(row, col);

    // First try exact cell match
    if let Some(props) = metadata.get(&cell_addr) {
        if let Some(role_str) = props.get("role") {
            if let Some(role) = Role::from_str(role_str) {
                return Some(role);
            }
        }
    }

    // Then try range matches (more expensive)
    for (target, props) in metadata.iter() {
        if target.contains(':') {
            // It's a range like "A3:I3"
            if cell_in_range(&cell_addr, row, col, target) {
                if let Some(role_str) = props.get("role") {
                    if let Some(role) = Role::from_str(role_str) {
                        return Some(role);
                    }
                }
            }
        }
    }

    None
}

/// Format cell address from row/col (0-indexed)
fn format_cell_address(row: usize, col: usize) -> String {
    let col_str = col_to_letter(col);
    format!("{}{}", col_str, row + 1)
}

/// Convert column index to letter(s): 0->A, 1->B, 26->AA
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

/// Check if a cell is within a range like "A3:I3"
fn cell_in_range(cell_addr: &str, row: usize, col: usize, range: &str) -> bool {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() != 2 {
        return false;
    }

    let (start_row, start_col) = match parse_cell_ref(parts[0]) {
        Some(rc) => rc,
        None => return false,
    };
    let (end_row, end_col) = match parse_cell_ref(parts[1]) {
        Some(rc) => rc,
        None => return false,
    };

    row >= start_row && row <= end_row && col >= start_col && col <= end_col
}

/// Parse cell reference like "A3" into (row, col) 0-indexed
fn parse_cell_ref(cell: &str) -> Option<(usize, usize)> {
    let cell = cell.trim();
    if cell.is_empty() {
        return None;
    }

    let mut col_end = 0;
    for (i, c) in cell.chars().enumerate() {
        if c.is_ascii_alphabetic() {
            col_end = i + 1;
        } else {
            break;
        }
    }

    if col_end == 0 || col_end >= cell.len() {
        return None;
    }

    let col_str = &cell[..col_end];
    let row_str = &cell[col_end..];

    let col = letter_to_col(col_str)?;
    let row: usize = row_str.parse().ok()?;

    if row == 0 {
        return None; // Row is 1-indexed in cell refs
    }

    Some((row - 1, col))
}

/// Convert column letter(s) to index: A->0, B->1, AA->26
fn letter_to_col(s: &str) -> Option<usize> {
    let mut col = 0usize;
    for c in s.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    if col == 0 {
        None
    } else {
        Some(col - 1)
    }
}

/// Format a number for display based on role format
pub fn format_number_for_display(value: f64, format: NumberDisplayFormat) -> String {
    match format {
        NumberDisplayFormat::Currency => {
            if value < 0.0 {
                format!("-${}", format_with_commas(value.abs(), 2))
            } else {
                format!("${}", format_with_commas(value, 2))
            }
        }
        NumberDisplayFormat::Percent => {
            format!("{:.2}%", value * 100.0)
        }
        NumberDisplayFormat::Integer => {
            format_with_commas(value.round(), 0)
        }
    }
}

/// Format number with thousand separators
fn format_with_commas(value: f64, decimals: usize) -> String {
    let formatted = if decimals == 0 {
        format!("{:.0}", value)
    } else {
        format!("{:.1$}", value, decimals)
    };

    // Split into integer and decimal parts
    let parts: Vec<&str> = formatted.split('.').collect();
    let int_part = parts[0];
    let dec_part = parts.get(1);

    // Add commas to integer part (comma goes after digit in reversed form
    // so it appears before the group when un-reversed)
    let int_with_commas: String = int_part
        .chars()
        .rev()
        .enumerate()
        .map(|(i, c)| {
            if i > 0 && i % 3 == 0 && c != '-' {
                format!("{},", c)
            } else {
                c.to_string()
            }
        })
        .collect::<Vec<_>>()
        .into_iter()
        .rev()
        .collect();

    match dec_part {
        Some(d) => format!("{}.{}", int_with_commas, d),
        None => int_with_commas,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
    }

    #[test]
    fn test_letter_to_col() {
        assert_eq!(letter_to_col("A"), Some(0));
        assert_eq!(letter_to_col("B"), Some(1));
        assert_eq!(letter_to_col("Z"), Some(25));
        assert_eq!(letter_to_col("AA"), Some(26));
        assert_eq!(letter_to_col("AB"), Some(27));
    }

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((0, 0)));
        assert_eq!(parse_cell_ref("B2"), Some((1, 1)));
        assert_eq!(parse_cell_ref("AA10"), Some((9, 26)));
    }

    #[test]
    fn test_cell_in_range() {
        assert!(cell_in_range("A3", 2, 0, "A3:I3"));
        assert!(cell_in_range("E3", 2, 4, "A3:I3"));
        assert!(cell_in_range("I3", 2, 8, "A3:I3"));
        assert!(!cell_in_range("A4", 3, 0, "A3:I3"));
        assert!(!cell_in_range("J3", 2, 9, "A3:I3"));
    }

    #[test]
    fn test_format_currency() {
        assert_eq!(format_number_for_display(1234.56, NumberDisplayFormat::Currency), "$1,234.56");
        assert_eq!(format_number_for_display(-1234.56, NumberDisplayFormat::Currency), "-$1,234.56");
        assert_eq!(format_number_for_display(1000000.0, NumberDisplayFormat::Currency), "$1,000,000.00");
    }

    #[test]
    fn test_format_percent() {
        assert_eq!(format_number_for_display(0.035, NumberDisplayFormat::Percent), "3.50%");
        assert_eq!(format_number_for_display(0.5, NumberDisplayFormat::Percent), "50.00%");
    }

    #[test]
    fn test_format_integer() {
        assert_eq!(format_number_for_display(1234.56, NumberDisplayFormat::Integer), "1,235");
        assert_eq!(format_number_for_display(1000000.0, NumberDisplayFormat::Integer), "1,000,000");
    }
}
