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
    Time,      // HH:MM:SS (fractional day → time of day)
    DateTime,  // Date + Time combined
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
// Excel has a famous bug: it treats 1900 as a leap year (it wasn't).
// Serial 60 = Feb 29, 1900 in Excel, which doesn't exist in reality.
// We replicate this bug for Excel compatibility.

/// Convert year/month/day to Excel serial date number (1900 date system)
/// Replicates Excel's 1900 leap year bug for compatibility.
pub fn date_to_serial(year: i32, month: u32, day: u32) -> f64 {
    // Handle the fake Feb 29, 1900
    if year == 1900 && month == 2 && day == 29 {
        return 60.0;
    }

    // Calculate days from 1900-01-01 (serial 1)
    let mut serial: i64 = 0;

    // Add days for complete years from 1900 to year-1
    for y in 1900..year {
        serial += if is_leap_year(y) { 366 } else { 365 };
    }

    // Add days for complete months in the current year
    let days_in_month = days_in_month_for_year(year);
    for m in 1..month {
        serial += days_in_month[(m - 1) as usize] as i64;
    }

    // Add days in current month
    serial += day as i64;

    // Excel's bug: dates on or after March 1, 1900 are off by 1
    // because Excel thinks Feb 29, 1900 existed (serial 60)
    // So we add 1 to account for the fake leap day
    if year > 1900 || (year == 1900 && month >= 3) {
        serial += 1;
    }

    serial as f64
}

/// Convert Excel serial date number to (year, month, day) (1900 date system)
/// Handles Excel's 1900 leap year bug for compatibility.
pub fn serial_to_date(serial: f64) -> (i32, u32, u32) {
    let serial = serial.floor() as i64;

    // Handle special cases
    if serial < 1 {
        return (1900, 1, 1);
    }

    // Excel's fake Feb 29, 1900
    if serial == 60 {
        return (1900, 2, 29);
    }

    // For serials > 60, we need to subtract 1 to account for the fake leap day
    // before doing the real calendar math
    let adjusted_serial = if serial > 60 { serial - 1 } else { serial };

    // Now convert using correct calendar (where 1900 is NOT a leap year)
    let mut remaining = adjusted_serial - 1; // serial 1 = Jan 1, 1900
    let mut year = 1900i32;

    // Find the year
    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if remaining < days_in_year {
            break;
        }
        remaining -= days_in_year;
        year += 1;
    }

    // Find the month
    let days_in_month = days_in_month_for_year(year);
    let mut month = 1u32;
    for &days in &days_in_month {
        if remaining < days as i64 {
            break;
        }
        remaining -= days as i64;
        month += 1;
    }

    let day = (remaining + 1) as u32;
    (year, month, day)
}

/// Check if a year is a leap year (correct Gregorian calendar)
fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

/// Get days in each month for a given year
fn days_in_month_for_year(year: i32) -> [u8; 12] {
    if is_leap_year(year) {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    }
}

/// Convert from 1904 date system serial to 1900 date system serial
/// Mac Excel uses 1904 system; Windows Excel uses 1900 system.
/// The difference is 1462 days.
pub fn serial_1904_to_1900(serial_1904: f64) -> f64 {
    // 1904 system has epoch Jan 1, 1904
    // 1900 system has epoch Jan 1, 1900 (with the fake Feb 29)
    // The difference is 1462 days
    serial_1904 + 1462.0
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

/// Format the time portion of a serial (fractional part)
/// Time is stored as fraction of a day: 0.5 = 12:00:00, 0.25 = 6:00:00
pub fn format_time(serial: f64) -> String {
    let fraction = serial.fract().abs();
    let total_seconds = (fraction * 86400.0).round() as u32; // 86400 = 24*60*60

    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;

    format!("{:02}:{:02}:{:02}", hours, minutes, seconds)
}

/// Convert time components to serial fraction
pub fn time_to_serial(hours: u32, minutes: u32, seconds: u32) -> f64 {
    let total_seconds = hours * 3600 + minutes * 60 + seconds;
    total_seconds as f64 / 86400.0
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
            NumberFormat::Time => {
                format_time(n)
            }
            NumberFormat::DateTime => {
                let date_part = format_date(n, DateStyle::Short);
                let time_part = format_time(n);
                format!("{} {}", date_part, time_part)
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

    // Excel date system tests (1900 system with leap year bug)

    #[test]
    fn test_excel_serial_basic_dates() {
        // Excel serial 1 = Jan 1, 1900
        assert_eq!(serial_to_date(1.0), (1900, 1, 1));
        // Excel serial 2 = Jan 2, 1900
        assert_eq!(serial_to_date(2.0), (1900, 1, 2));
        // Jan 31, 1900 = serial 31
        assert_eq!(serial_to_date(31.0), (1900, 1, 31));
        // Feb 1, 1900 = serial 32
        assert_eq!(serial_to_date(32.0), (1900, 2, 1));
    }

    #[test]
    fn test_excel_leap_year_bug() {
        // Excel's famous bug: Feb 29, 1900 = serial 60 (doesn't exist in reality)
        assert_eq!(serial_to_date(60.0), (1900, 2, 29));
        // Feb 28, 1900 = serial 59
        assert_eq!(serial_to_date(59.0), (1900, 2, 28));
        // Mar 1, 1900 = serial 61
        assert_eq!(serial_to_date(61.0), (1900, 3, 1));
    }

    #[test]
    fn test_excel_serial_roundtrip_around_bug() {
        // Test roundtrip for dates around the leap year bug
        assert_eq!(date_to_serial(1900, 2, 28), 59.0);
        assert_eq!(date_to_serial(1900, 2, 29), 60.0); // Fake date
        assert_eq!(date_to_serial(1900, 3, 1), 61.0);

        // Verify roundtrip
        assert_eq!(serial_to_date(date_to_serial(1900, 2, 28)), (1900, 2, 28));
        assert_eq!(serial_to_date(date_to_serial(1900, 2, 29)), (1900, 2, 29));
        assert_eq!(serial_to_date(date_to_serial(1900, 3, 1)), (1900, 3, 1));
    }

    #[test]
    fn test_excel_serial_modern_dates() {
        // Known Excel values for modern dates
        // Jan 1, 2000 = serial 36526
        assert_eq!(serial_to_date(36526.0), (2000, 1, 1));
        assert_eq!(date_to_serial(2000, 1, 1), 36526.0);

        // Jan 1, 2024 = serial 45292
        assert_eq!(serial_to_date(45292.0), (2024, 1, 1));
        assert_eq!(date_to_serial(2024, 1, 1), 45292.0);

        // Dec 31, 2024 = serial 45657
        assert_eq!(serial_to_date(45657.0), (2024, 12, 31));
        assert_eq!(date_to_serial(2024, 12, 31), 45657.0);
    }

    #[test]
    fn test_excel_serial_leap_years() {
        // Feb 29, 2000 (real leap year) = serial 36585
        assert_eq!(serial_to_date(36585.0), (2000, 2, 29));
        assert_eq!(date_to_serial(2000, 2, 29), 36585.0);

        // Feb 29, 2024 (real leap year) = serial 45351
        assert_eq!(serial_to_date(45351.0), (2024, 2, 29));
        assert_eq!(date_to_serial(2024, 2, 29), 45351.0);
    }

    #[test]
    fn test_excel_1904_to_1900_conversion() {
        // 1904 system epoch is Jan 1, 1904
        // In 1904 system, Jan 1, 1904 = serial 0
        // In 1900 system, Jan 1, 1904 = serial 1462
        assert_eq!(serial_1904_to_1900(0.0), 1462.0);

        // Verify the converted date
        assert_eq!(serial_to_date(serial_1904_to_1900(0.0)), (1904, 1, 1));

        // Jan 2, 1904 in 1904 system = serial 1
        // Converted to 1900 system = serial 1463
        assert_eq!(serial_1904_to_1900(1.0), 1463.0);
        assert_eq!(serial_to_date(1463.0), (1904, 1, 2));
    }

    #[test]
    fn test_excel_time_fraction() {
        // Time is the fractional part of the serial
        // 0.5 = noon (12:00:00)
        // 0.25 = 6:00 AM
        // 0.75 = 6:00 PM
        let serial = 45292.5; // Jan 1, 2024 at noon
        let (year, month, day) = serial_to_date(serial);
        assert_eq!((year, month, day), (2024, 1, 1));
        // Fractional part preserved
        assert!((serial.fract() - 0.5).abs() < 0.0001);
    }
}
