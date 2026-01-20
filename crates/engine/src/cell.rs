use serde::{Deserialize, Serialize};

use super::formula::parser::{self, Expr};

/// Horizontal text alignment
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum Alignment {
    #[default]
    Left,
    Center,
    Right,
}

/// Vertical text alignment
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum VerticalAlignment {
    Top,
    #[default]
    Middle,
    Bottom,
}

/// Text overflow behavior
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum TextOverflow {
    #[default]
    Clip,       // Text is clipped at cell boundary
    Wrap,       // Text wraps to multiple lines within the cell
    Overflow,   // Text overflows into adjacent empty cells
}

/// Date format style
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum DateStyle {
    #[default]
    Short,      // 1/18/2026
    Long,       // January 18, 2026
    Iso,        // 2026-01-18
}

/// Number format type
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum NumberFormat {
    #[default]
    General,
    Number { decimals: u8 },
    Currency { decimals: u8 },
    Percent { decimals: u8 },
    Date { style: DateStyle },
}

/// Cell formatting options
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
pub struct CellFormat {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub alignment: Alignment,
    pub vertical_alignment: VerticalAlignment,
    pub text_overflow: TextOverflow,
    pub number_format: NumberFormat,
    pub font_family: Option<String>,  // None = inherit from settings
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    #[serde(skip)]
    Formula { source: String, ast: Option<Expr> },
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

// Excel serial date epoch: December 30, 1899 (day 0)
// This matches Excel's date system (with the intentional 1900 leap year bug compatibility)

/// Convert year/month/day to Excel serial date number
pub fn date_to_serial(year: i32, month: u32, day: u32) -> f64 {
    // Days from Excel epoch (1899-12-30) to Unix epoch (1970-01-01) is 25569
    const EXCEL_EPOCH_OFFSET: i64 = 25569;

    // Use a simple calculation for dates
    // Days in each month (non-leap year)
    let days_in_month = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    // Calculate days from year 1 to start of this year
    let y = year as i64 - 1;
    let mut days = y * 365 + y / 4 - y / 100 + y / 400;

    // Add days for months in this year
    for m in 1..month {
        days += days_in_month[m as usize] as i64;
    }

    // Add leap day if applicable
    let is_leap = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
    if is_leap && month > 2 {
        days += 1;
    }

    // Add days in current month
    days += day as i64;

    // Convert to Excel serial (days since 1899-12-30)
    // Days from year 1 to 1899-12-30
    let epoch_days = {
        let y = 1899_i64;
        let base = y * 365 + y / 4 - y / 100 + y / 400;
        base + 334 + 30  // Dec 30 = 334 days into year + 30 for December
    };

    (days - epoch_days) as f64
}

/// Convert Excel serial date number to (year, month, day)
pub fn serial_to_date(serial: f64) -> (i32, u32, u32) {
    // Excel serial 1 = January 1, 1900
    // We need to convert back
    let serial = serial.floor() as i64;

    // Days from Excel epoch to calculate
    // Excel epoch is 1899-12-30, so serial 1 = 1899-12-31, serial 2 = 1900-01-01
    // Actually Excel considers serial 1 = 1900-01-01

    // Simplified: add serial days to 1899-12-30
    let mut days = serial;

    // Start from 1900-01-01 (serial = 1)
    let mut year = 1900;
    let mut remaining = days - 1;  // serial 1 = day 0 of our count

    loop {
        let is_leap = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
        let days_in_year = if is_leap { 366 } else { 365 };

        if remaining < days_in_year {
            break;
        }
        remaining -= days_in_year;
        year += 1;
    }

    // Now find month
    let is_leap = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
    let days_in_month = if is_leap {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month = 1u32;
    for days in days_in_month.iter() {
        if remaining < *days {
            break;
        }
        remaining -= *days;
        month += 1;
    }

    let day = (remaining + 1) as u32;

    (year, month, day)
}

/// Format a serial date according to style
pub fn format_date(serial: f64, style: DateStyle) -> String {
    let (year, month, day) = serial_to_date(serial);

    match style {
        DateStyle::Short => format!("{}/{}/{}", month, day, year),
        DateStyle::Long => {
            let month_name = match month {
                1 => "January", 2 => "February", 3 => "March", 4 => "April",
                5 => "May", 6 => "June", 7 => "July", 8 => "August",
                9 => "September", 10 => "October", 11 => "November", 12 => "December",
                _ => "Unknown",
            };
            format!("{} {}, {}", month_name, day, year)
        }
        DateStyle::Iso => format!("{:04}-{:02}-{:02}", year, month, day),
    }
}

/// Try to parse a date string, returns serial number if successful
pub fn parse_date(input: &str) -> Option<f64> {
    let trimmed = input.trim();

    // Try ISO format: YYYY-MM-DD
    if let Some((year, rest)) = trimmed.split_once('-') {
        if let Some((month, day)) = rest.split_once('-') {
            if let (Ok(y), Ok(m), Ok(d)) = (year.parse::<i32>(), month.parse::<u32>(), day.parse::<u32>()) {
                if m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1900 && y <= 9999 {
                    return Some(date_to_serial(y, m, d));
                }
            }
        }
    }

    // Try US format: M/D/YYYY or MM/DD/YYYY
    if let Some((month, rest)) = trimmed.split_once('/') {
        if let Some((day, year)) = rest.split_once('/') {
            if let (Ok(m), Ok(d), Ok(y)) = (month.parse::<u32>(), day.parse::<u32>(), year.parse::<i32>()) {
                if m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1900 && y <= 9999 {
                    return Some(date_to_serial(y, m, d));
                }
            }
        }
    }

    None
}

impl CellValue {
    pub fn from_input(input: &str) -> Self {
        let trimmed = input.trim();

        if trimmed.is_empty() {
            return CellValue::Empty;
        }

        if trimmed.starts_with('=') {
            let ast = parser::parse(trimmed).ok();
            return CellValue::Formula {
                source: trimmed.to_string(),
                ast,
            };
        }

        if let Ok(num) = trimmed.parse::<f64>() {
            return CellValue::Number(num);
        }

        CellValue::Text(trimmed.to_string())
    }

    pub fn raw_display(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => {
                if n.fract() == 0.0 {
                    format!("{}", *n as i64)
                } else {
                    format!("{:.2}", n)
                }
            }
            CellValue::Formula { source, .. } => source.clone(),
        }
    }

    /// Format a number according to the specified format
    pub fn format_number(n: f64, format: &NumberFormat) -> String {
        match format {
            NumberFormat::General => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", n as i64)
                } else {
                    format!("{:.2}", n)
                }
            }
            NumberFormat::Number { decimals } => {
                format!("{:.*}", *decimals as usize, n)
            }
            NumberFormat::Currency { decimals } => {
                if n < 0.0 {
                    format!("-${:.*}", *decimals as usize, n.abs())
                } else {
                    format!("${:.*}", *decimals as usize, n)
                }
            }
            NumberFormat::Percent { decimals } => {
                format!("{:.*}%", *decimals as usize, n * 100.0)
            }
            NumberFormat::Date { style } => {
                format_date(n, *style)
            }
        }
    }

    /// Display value with formatting applied
    pub fn formatted_display(&self, format: &CellFormat) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => Self::format_number(*n, &format.number_format),
            CellValue::Formula { source, .. } => source.clone(),
        }
    }

    pub fn as_number(&self) -> f64 {
        match self {
            CellValue::Number(n) => *n,
            CellValue::Text(s) => s.parse().unwrap_or(0.0),
            _ => 0.0,
        }
    }
}

/// Spill information for a cell that is the source of array spill
#[derive(Debug, Clone, Default)]
pub struct SpillInfo {
    /// The dimensions of the spilled array (rows, cols)
    pub rows: usize,
    pub cols: usize,
}

/// Spill error state - tracks when array formula spill is blocked
#[derive(Debug, Clone, Default)]
pub struct SpillError {
    /// The cell blocking the spill
    pub blocked_by: (usize, usize),
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Cell {
    pub value: CellValue,
    pub format: CellFormat,
    /// If this cell receives spill data, points to the parent formula cell (row, col)
    #[serde(skip)]
    pub spill_parent: Option<(usize, usize)>,
    /// If this cell has a formula that produces an array, contains spill info
    #[serde(skip)]
    pub spill_info: Option<SpillInfo>,
    /// If this cell has an array formula that can't spill, contains error info
    #[serde(skip)]
    pub spill_error: Option<SpillError>,
}

impl Cell {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn set(&mut self, input: &str) {
        self.value = CellValue::from_input(input);
        // Clear spill state when cell is edited
        self.spill_parent = None;
        self.spill_info = None;
        self.spill_error = None;
    }

    /// Check if this cell is receiving spill data from another cell
    pub fn is_spill_receiver(&self) -> bool {
        self.spill_parent.is_some()
    }

    /// Check if this cell is a spill parent (has array formula that spills)
    pub fn is_spill_parent(&self) -> bool {
        self.spill_info.is_some()
    }

    /// Check if this cell has a blocked spill error
    pub fn has_spill_error(&self) -> bool {
        self.spill_error.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_overflow_default_is_clip() {
        let format = CellFormat::default();
        assert_eq!(format.text_overflow, TextOverflow::Clip);
    }

    #[test]
    fn test_vertical_alignment_default_is_middle() {
        let format = CellFormat::default();
        assert_eq!(format.vertical_alignment, VerticalAlignment::Middle);
    }

    #[test]
    fn test_cell_format_defaults() {
        let format = CellFormat::default();
        assert!(!format.bold);
        assert!(!format.italic);
        assert!(!format.underline);
        assert!(!format.strikethrough);
        assert_eq!(format.alignment, Alignment::Left);
        assert_eq!(format.vertical_alignment, VerticalAlignment::Middle);
        assert_eq!(format.text_overflow, TextOverflow::Clip);
        assert_eq!(format.number_format, NumberFormat::General);
    }

    #[test]
    fn test_text_overflow_equality() {
        assert_eq!(TextOverflow::Clip, TextOverflow::Clip);
        assert_eq!(TextOverflow::Wrap, TextOverflow::Wrap);
        assert_eq!(TextOverflow::Overflow, TextOverflow::Overflow);
        assert_ne!(TextOverflow::Clip, TextOverflow::Wrap);
        assert_ne!(TextOverflow::Clip, TextOverflow::Overflow);
        assert_ne!(TextOverflow::Wrap, TextOverflow::Overflow);
    }

    #[test]
    fn test_vertical_alignment_equality() {
        assert_eq!(VerticalAlignment::Top, VerticalAlignment::Top);
        assert_eq!(VerticalAlignment::Middle, VerticalAlignment::Middle);
        assert_eq!(VerticalAlignment::Bottom, VerticalAlignment::Bottom);
        assert_ne!(VerticalAlignment::Top, VerticalAlignment::Middle);
        assert_ne!(VerticalAlignment::Top, VerticalAlignment::Bottom);
        assert_ne!(VerticalAlignment::Middle, VerticalAlignment::Bottom);
    }
}
