use serde::{Deserialize, Serialize};

use super::formula::parser::{self, ParsedExpr};

/// Horizontal text alignment
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum Alignment {
    #[default]
    General,  // Auto: numbers right-align, text left-aligns (Excel default)
    Left,
    Center,
    Right,
    /// Center text across selected columns (Excel: Center Across Selection)
    CenterAcrossSelection,
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

/// Negative number display style
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum NegativeStyle {
    #[default]
    Minus,       // -1,234.56
    Parens,      // (1,234.56)
    RedMinus,    // -1,234.56 in red
    RedParens,   // (1,234.56) in red
}

impl NegativeStyle {
    /// True if this style should render the value in red
    pub fn is_red(&self) -> bool {
        matches!(self, NegativeStyle::RedMinus | NegativeStyle::RedParens)
    }

    /// True if this style uses parentheses instead of minus sign
    pub fn uses_parens(&self) -> bool {
        matches!(self, NegativeStyle::Parens | NegativeStyle::RedParens)
    }

    /// Convert to integer for storage (0-3)
    pub fn to_int(&self) -> i32 {
        match self {
            NegativeStyle::Minus => 0,
            NegativeStyle::Parens => 1,
            NegativeStyle::RedMinus => 2,
            NegativeStyle::RedParens => 3,
        }
    }

    /// Convert from integer (0-3), defaults to Minus for unknown values
    pub fn from_int(i: i32) -> Self {
        match i {
            1 => NegativeStyle::Parens,
            2 => NegativeStyle::RedMinus,
            3 => NegativeStyle::RedParens,
            _ => NegativeStyle::Minus,
        }
    }
}

/// Number format type
#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
pub enum NumberFormat {
    #[default]
    General,
    Number {
        decimals: u8,
        #[serde(default)]
        thousands: bool,
        #[serde(default)]
        negative: NegativeStyle,
    },
    Currency {
        decimals: u8,
        #[serde(default)]
        thousands: bool,
        #[serde(default)]
        negative: NegativeStyle,
        #[serde(default)]
        symbol: Option<String>,
    },
    Percent { decimals: u8 },
    Date { style: DateStyle },
    Time,      // HH:MM:SS (fractional day → time of day)
    DateTime,  // Date + Time combined
    /// Raw Excel format code (e.g. "#,##0.00", "$#,##0")
    Custom(String),
}

impl NumberFormat {
    /// UI default Number format: thousands separator on, negative minus
    pub fn number(decimals: u8) -> Self {
        NumberFormat::Number {
            decimals: decimals.min(10),
            thousands: true,
            negative: NegativeStyle::Minus,
        }
    }

    /// UI default Currency format: thousands separator on, negative parens, default $ symbol
    pub fn currency(decimals: u8) -> Self {
        NumberFormat::Currency {
            decimals: decimals.min(10),
            thousands: true,
            negative: NegativeStyle::Parens,
            symbol: None,
        }
    }

    /// Backward-compatible Number format (decoder default): no thousands, negative minus
    pub fn number_compat(decimals: u8) -> Self {
        NumberFormat::Number {
            decimals: decimals.min(10),
            thousands: false,
            negative: NegativeStyle::Minus,
        }
    }

    /// Backward-compatible Currency format (decoder default): no thousands, negative minus
    pub fn currency_compat(decimals: u8) -> Self {
        NumberFormat::Currency {
            decimals: decimals.min(10),
            thousands: false,
            negative: NegativeStyle::Minus,
            symbol: None,
        }
    }

    /// Returns decimals for Number/Currency/Percent, None for others
    pub fn decimals(&self) -> Option<u8> {
        match self {
            NumberFormat::Number { decimals, .. } => Some(*decimals),
            NumberFormat::Currency { decimals, .. } => Some(*decimals),
            NumberFormat::Percent { decimals } => Some(*decimals),
            _ => None,
        }
    }

    /// True when value < 0 and the negative style includes red coloring
    pub fn should_render_red(&self, value: f64) -> bool {
        if value >= 0.0 {
            return false;
        }
        match self {
            NumberFormat::Number { negative, .. } => negative.is_red(),
            NumberFormat::Currency { negative, .. } => negative.is_red(),
            _ => false,
        }
    }
}

/// Border style (line thickness)
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq, Eq, Hash)]
pub enum BorderStyle {
    #[default]
    None,
    Thin,      // 1px
    Medium,    // 2px
    Thick,     // 3px
}

impl BorderStyle {
    /// Numeric weight for precedence comparison: None(0) < Thin(1) < Medium(2) < Thick(3)
    pub fn weight(self) -> u8 {
        match self {
            BorderStyle::None => 0,
            BorderStyle::Thin => 1,
            BorderStyle::Medium => 2,
            BorderStyle::Thick => 3,
        }
    }
}

/// Border specification for a single cell edge
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq, Eq, Hash)]
pub struct CellBorder {
    pub style: BorderStyle,
    /// RGBA color. None = black. Stored but normalized to black in v0.2.2.
    pub color: Option<[u8; 4]>,
}

impl CellBorder {
    /// Returns true if this border should be rendered (has a visible style)
    pub fn is_set(&self) -> bool {
        self.style != BorderStyle::None
    }

    /// Creates a thin black border (the only style available in v0.2.2 UI)
    pub fn thin() -> Self {
        Self {
            style: BorderStyle::Thin,
            color: None,
        }
    }
}

/// Returns the winning border for a shared edge (precedence logic).
/// border_a takes precedence if set, otherwise border_b.
pub fn effective_border(border_a: CellBorder, border_b: CellBorder) -> CellBorder {
    if border_a.style != BorderStyle::None {
        border_a
    } else {
        border_b
    }
}

/// Returns the border with the strongest style (None < Thin < Medium < Thick).
/// Used for merged cell edge resolution: scan edge cells, keep the "thickest" border.
pub fn max_border(a: CellBorder, b: CellBorder) -> CellBorder {
    if a.style.weight() >= b.style.weight() { a } else { b }
}

/// Returns the border to actually draw (v0.2.2: normalize to Thin/black).
/// Medium/Thick are stored but rendered as Thin to avoid overlay complexity.
pub fn render_border(border: CellBorder) -> CellBorder {
    if border.style != BorderStyle::None {
        CellBorder {
            style: BorderStyle::Thin,
            color: None,
        }
    } else {
        border
    }
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
    /// Font size in points. None = inherit from settings.
    #[serde(default)]
    pub font_size: Option<f32>,
    /// Font color as RGBA. None = inherit from theme.
    #[serde(default)]
    pub font_color: Option<[u8; 4]>,
    /// Background fill color as RGBA. None = transparent/default.
    #[serde(default)]
    pub background_color: Option<[u8; 4]>,
    /// Top edge border
    #[serde(default)]
    pub border_top: CellBorder,
    /// Right edge border
    #[serde(default)]
    pub border_right: CellBorder,
    /// Bottom edge border
    #[serde(default)]
    pub border_bottom: CellBorder,
    /// Left edge border
    #[serde(default)]
    pub border_left: CellBorder,
}

impl CellFormat {
    /// Returns true if any border edge is set (non-None style).
    pub fn has_any_border(&self) -> bool {
        self.border_top.is_set() || self.border_right.is_set()
            || self.border_bottom.is_set() || self.border_left.is_set()
    }

    /// Merge a CellFormatOverride on top of this base format.
    /// Override fields that are `Some` replace the base; `None` fields keep the base value.
    pub fn merge_override(&self, ovr: &CellFormatOverride) -> CellFormat {
        CellFormat {
            bold: ovr.bold.unwrap_or(self.bold),
            italic: ovr.italic.unwrap_or(self.italic),
            underline: ovr.underline.unwrap_or(self.underline),
            strikethrough: ovr.strikethrough.unwrap_or(self.strikethrough),
            alignment: ovr.alignment.unwrap_or(self.alignment),
            vertical_alignment: ovr.vertical_alignment.unwrap_or(self.vertical_alignment),
            text_overflow: ovr.text_overflow.unwrap_or(self.text_overflow),
            number_format: ovr.number_format.clone().unwrap_or_else(|| self.number_format.clone()),
            font_family: ovr.font_family.clone().unwrap_or_else(|| self.font_family.clone()),
            font_size: ovr.font_size.unwrap_or(self.font_size),
            font_color: ovr.font_color.unwrap_or(self.font_color),
            background_color: ovr.background_color.unwrap_or(self.background_color),
            border_top: ovr.border_top.unwrap_or(self.border_top),
            border_right: ovr.border_right.unwrap_or(self.border_right),
            border_bottom: ovr.border_bottom.unwrap_or(self.border_bottom),
            border_left: ovr.border_left.unwrap_or(self.border_left),
        }
    }
}

/// Partial format override with all-Option fields.
/// Used during XLSX import to represent style deltas.
/// `None` = "not overridden, use base style"; `Some(v)` = "explicitly set to v".
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CellFormatOverride {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_alignment: Option<VerticalAlignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_overflow: Option<TextOverflow>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<NumberFormat>,
    /// None = not overridden; Some(None) = explicitly inherit; Some(Some(s)) = explicitly set
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_family: Option<Option<String>>,
    /// None = not overridden; Some(None) = explicitly inherit; Some(Some(v)) = explicitly set
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size: Option<Option<f32>>,
    /// None = not overridden; Some(None) = explicitly inherit; Some(Some(v)) = explicitly set
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<Option<[u8; 4]>>,
    /// None = not overridden; Some(None) = explicitly inherit; Some(Some(v)) = explicitly set
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background_color: Option<Option<[u8; 4]>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_top: Option<CellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_right: Option<CellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_bottom: Option<CellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_left: Option<CellBorder>,
}

impl CellFormatOverride {
    /// Convert a full CellFormat into an override where all fields are `Some`.
    pub fn from_format(format: &CellFormat) -> Self {
        Self {
            bold: Some(format.bold),
            italic: Some(format.italic),
            underline: Some(format.underline),
            strikethrough: Some(format.strikethrough),
            alignment: Some(format.alignment),
            vertical_alignment: Some(format.vertical_alignment),
            text_overflow: Some(format.text_overflow),
            number_format: Some(format.number_format.clone()),
            font_family: Some(format.font_family.clone()),
            font_size: Some(format.font_size),
            font_color: Some(format.font_color),
            background_color: Some(format.background_color),
            border_top: Some(format.border_top),
            border_right: Some(format.border_right),
            border_bottom: Some(format.border_bottom),
            border_left: Some(format.border_left),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    #[serde(skip)]
    Formula { source: String, ast: Option<ParsedExpr> },
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

/// Format a number using a raw Excel format code string.
///
/// Covers common finance patterns:
/// - `#,##0` / `#,##0.00` → thousands separator with decimals
/// - `(#,##0.00)` → negative parentheses
/// - `$#,##0` / `$#,##0.00` → dollar + thousands
/// - `0%` / `0.00%` → percent
/// - Unknown codes → plain formatted number
pub fn format_with_custom_code(n: f64, code: &str) -> String {
    // Step 1: Strip quoted literal text from the format code.
    // Excel wraps literal characters in double quotes: "$"#,##0 or #,##0.00" USD"
    // Also handle backslash-escaped single characters: \$ → $
    // We extract prefix/suffix literals and a clean numeric pattern.
    let clean_code = strip_format_quotes(code);

    // Strip accounting padding characters (_ and *) and surrounding whitespace
    // _ followed by a char means "leave space equal to that char's width" → strip both
    // * followed by a char means "repeat that char to fill" → strip both
    let clean = strip_accounting_padding(&clean_code);

    // Check for multi-section format (positive;negative;zero)
    // Must split carefully: semicolons inside quotes were already resolved above
    let section = if clean.contains(';') {
        let sections: Vec<&str> = clean.split(';').collect();
        if n < 0.0 && sections.len() >= 2 {
            sections[1].trim()
        } else if n == 0.0 && sections.len() >= 3 {
            sections[2].trim()
        } else {
            sections[0].trim()
        }
    } else {
        clean.as_str()
    };

    // Detect negative parentheses pattern: e.g. "(#,##0.00)" or "($#,##0.00)"
    let (use_parens, inner) = if section.starts_with('(') && section.ends_with(')') {
        (true, &section[1..section.len() - 1])
    } else {
        (false, section)
    };

    // Split into: prefix literals, numeric pattern, suffix literals
    // e.g. "$#,##0.00 USD" → prefix="$", pattern="#,##0.00", suffix=" USD"
    let (prefix, pattern, suffix) = split_format_parts(inner);

    // Detect percent suffix (in the numeric pattern or suffix)
    let (is_percent, pattern, suffix) = if pattern.ends_with('%') {
        (true, &pattern[..pattern.len() - 1], suffix)
    } else if suffix.starts_with('%') {
        (true, pattern, &suffix[1..])
    } else {
        (false, pattern, suffix)
    };

    // If pattern has no number format characters (#, 0), treat as literal text
    if !pattern.contains('#') && !pattern.contains('0') {
        return section.to_string();
    }

    let value = if is_percent { n * 100.0 } else { n.abs() };

    // Count decimal places from pattern
    let decimals = if let Some(dot_pos) = pattern.find('.') {
        pattern[dot_pos + 1..].chars().take_while(|&c| c == '0' || c == '#').count()
    } else {
        0
    };

    // Check for thousands separator
    let use_thousands = pattern.contains(',');

    // Format the number
    let formatted_num = if use_thousands {
        format_with_thousands(value, decimals)
    } else {
        format!("{:.*}", decimals, value)
    };

    // Assemble result with prefix and suffix
    let pct = if is_percent { "%" } else { "" };
    let abs_result = format!("{}{}{}{}", prefix, formatted_num, pct, suffix);

    if use_parens && n < 0.0 {
        format!("({})", abs_result)
    } else if !use_parens && n < 0.0 && !is_percent {
        format!("-{}", abs_result)
    } else {
        abs_result
    }
}

/// Strip double-quoted literal text and backslash-escaped characters from a format code.
/// Replaces "text" with the literal text (no quotes) and \c with c.
fn strip_format_quotes(code: &str) -> String {
    let mut result = String::with_capacity(code.len());
    let mut chars = code.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                // Consume everything until closing quote
                while let Some(inner) = chars.next() {
                    if inner == '"' {
                        break;
                    }
                    result.push(inner);
                }
            }
            '\\' => {
                // Next character is a literal
                if let Some(escaped) = chars.next() {
                    result.push(escaped);
                }
            }
            _ => result.push(c),
        }
    }
    result
}

/// Strip accounting padding characters: _X (space for width of X) and *X (repeat X to fill).
/// Both consume the next character after them.
fn strip_accounting_padding(code: &str) -> String {
    let mut result = String::with_capacity(code.len());
    let mut chars = code.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '_' | '*' => {
                // Skip the next character (the padding/fill character)
                chars.next();
            }
            _ => result.push(c),
        }
    }
    result.trim().to_string()
}

/// Split a format section into prefix literals, numeric pattern, and suffix literals.
/// The numeric pattern is the portion containing #, 0, comma, dot.
/// Everything before the first format char is prefix; everything after the last is suffix.
/// e.g. "$#,##0.00 USD" → ("$", "#,##0.00", " USD")
fn split_format_parts(section: &str) -> (&str, &str, &str) {
    let format_chars = |c: char| matches!(c, '#' | '0' | ',' | '.');

    let first = section.find(format_chars);
    let last = section.rfind(format_chars);

    match (first, last) {
        (Some(f), Some(l)) => {
            let prefix = &section[..f];
            let pattern = &section[f..=l];
            let suffix = &section[l + 1..];
            (prefix, pattern, suffix)
        }
        _ => {
            // No format characters found — treat entire section as literal
            (section, "", "")
        }
    }
}

/// Format an absolute value with optional thousands grouping.
/// Works from the numeric value directly (no string parsing).
/// Decimals clamped to 0..=10 as a safety net against overflow.
fn format_grouped(abs: f64, decimals: u8, thousands: bool) -> String {
    let decimals = decimals.min(10) as u32;
    let scale = 10_i64.pow(decimals);
    let scaled = (abs * scale as f64).round();
    if !scaled.is_finite() {
        return abs.to_string();
    }
    let scaled = scaled as i64;
    let int_part = scaled / scale;
    let frac_part = (scaled % scale).abs();

    let int_str = if thousands {
        let raw = int_part.to_string();
        let mut result = String::with_capacity(raw.len() + raw.len() / 3);
        for (i, ch) in raw.chars().rev().enumerate() {
            if i > 0 && i % 3 == 0 {
                result.push(',');
            }
            result.push(ch);
        }
        result.chars().rev().collect()
    } else {
        int_part.to_string()
    };

    if decimals == 0 {
        int_str
    } else {
        format!("{}.{:0>width$}", int_str, frac_part, width = decimals as usize)
    }
}

/// Format a number with thousands separators
fn format_with_thousands(n: f64, decimals: usize) -> String {
    let abs = n.abs();
    let integer_part = abs.trunc() as u64;
    let int_str = integer_part.to_string();

    // Insert commas every 3 digits from the right
    let mut with_commas = String::with_capacity(int_str.len() + int_str.len() / 3);
    for (i, ch) in int_str.chars().enumerate() {
        if i > 0 && (int_str.len() - i) % 3 == 0 {
            with_commas.push(',');
        }
        with_commas.push(ch);
    }

    if decimals > 0 {
        let frac = abs.fract();
        let frac_str = format!("{:.*}", decimals, frac);
        // frac_str is like "0.12", take everything after the dot
        let dot_pos = frac_str.find('.').unwrap_or(1);
        format!("{}.{}", with_commas, &frac_str[dot_pos + 1..])
    } else {
        with_commas
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

    /// Check if this cell contains a cycle error (#CYCLE!).
    pub fn is_cycle_error(&self) -> bool {
        matches!(self, CellValue::Text(s) if s == "#CYCLE!")
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
            NumberFormat::Number { decimals, thousands, negative } => {
                let formatted = format_grouped(n.abs(), *decimals, *thousands);
                if n < 0.0 {
                    if negative.uses_parens() {
                        format!("({})", formatted)
                    } else {
                        format!("-{}", formatted)
                    }
                } else {
                    formatted
                }
            }
            NumberFormat::Currency { decimals, thousands, negative, symbol } => {
                let sym = symbol.as_deref().unwrap_or("$");
                let formatted = format_grouped(n.abs(), *decimals, *thousands);
                let prefixed = format!("{}{}", sym, formatted);
                if n < 0.0 {
                    if negative.uses_parens() {
                        format!("({})", prefixed)
                    } else {
                        format!("-{}", prefixed)
                    }
                } else {
                    prefixed
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
            NumberFormat::Custom(ref code) => {
                // Use ssfmt for full ECMA-376 format code rendering (accounting,
                // multi-section, date/time tokens, etc.). Falls back to our
                // simpler formatter if ssfmt can't parse the code.
                match ssfmt::format_default(n, code) {
                    Ok(s) => s,
                    Err(_) => format_with_custom_code(n, code),
                }
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

    /// Get the parsed AST for formula cells, if available.
    pub fn formula_ast(&self) -> Option<&parser::ParsedExpr> {
        match self {
            CellValue::Formula { ast, .. } => ast.as_ref(),
            _ => None,
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
    /// Index into workbook.style_table — tracks the base style from XLSX import.
    /// User edits modify `format` directly; this field preserves import provenance
    /// for future "reset to imported style" functionality.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<u32>,
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
        assert_eq!(format.alignment, Alignment::General);
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

    // Border tests

    #[test]
    fn test_border_style_default_is_none() {
        let style = BorderStyle::default();
        assert_eq!(style, BorderStyle::None);
    }

    #[test]
    fn test_cell_border_default_is_no_border() {
        let border = CellBorder::default();
        assert_eq!(border.style, BorderStyle::None);
        assert_eq!(border.color, None);
        assert!(!border.is_set());
    }

    #[test]
    fn test_cell_border_thin() {
        let border = CellBorder::thin();
        assert_eq!(border.style, BorderStyle::Thin);
        assert_eq!(border.color, None);
        assert!(border.is_set());
    }

    #[test]
    fn test_cell_format_border_defaults() {
        let format = CellFormat::default();
        assert!(!format.border_top.is_set());
        assert!(!format.border_right.is_set());
        assert!(!format.border_bottom.is_set());
        assert!(!format.border_left.is_set());
    }

    #[test]
    fn test_effective_border_precedence() {
        let thin = CellBorder::thin();
        let none = CellBorder::default();
        let medium = CellBorder { style: BorderStyle::Medium, color: None };

        // First border takes precedence if set
        assert_eq!(effective_border(thin, none), thin);
        assert_eq!(effective_border(thin, medium), thin);

        // Falls back to second if first is None
        assert_eq!(effective_border(none, thin), thin);
        assert_eq!(effective_border(none, medium), medium);

        // Both None returns None
        assert_eq!(effective_border(none, none), none);
    }

    #[test]
    fn test_max_border_picks_strongest() {
        let none = CellBorder::default();
        let thin = CellBorder::thin();
        let medium = CellBorder { style: BorderStyle::Medium, color: None };
        let thick = CellBorder { style: BorderStyle::Thick, color: None };

        // max_border always picks the stronger style
        assert_eq!(max_border(none, thin), thin);
        assert_eq!(max_border(thin, none), thin);
        assert_eq!(max_border(thin, medium), medium);
        assert_eq!(max_border(medium, thin), medium);
        assert_eq!(max_border(medium, thick), thick);
        assert_eq!(max_border(thick, medium), thick);
        assert_eq!(max_border(none, none), none);
        assert_eq!(max_border(thick, thick), thick);

        // tie-break: first wins (both equal weight)
        let thin_red = CellBorder { style: BorderStyle::Thin, color: Some([255, 0, 0, 255]) };
        assert_eq!(max_border(thin, thin_red), thin); // a wins on tie
    }

    #[test]
    fn test_border_style_weight_ordering() {
        assert!(BorderStyle::None.weight() < BorderStyle::Thin.weight());
        assert!(BorderStyle::Thin.weight() < BorderStyle::Medium.weight());
        assert!(BorderStyle::Medium.weight() < BorderStyle::Thick.weight());
    }

    #[test]
    fn test_render_border_normalizes_to_thin() {
        let thin = CellBorder::thin();
        let medium = CellBorder { style: BorderStyle::Medium, color: None };
        let thick = CellBorder { style: BorderStyle::Thick, color: None };
        let none = CellBorder::default();

        // All non-None styles render as Thin/black
        assert_eq!(render_border(thin), CellBorder { style: BorderStyle::Thin, color: None });
        assert_eq!(render_border(medium), CellBorder { style: BorderStyle::Thin, color: None });
        assert_eq!(render_border(thick), CellBorder { style: BorderStyle::Thin, color: None });

        // None stays None
        assert_eq!(render_border(none), CellBorder { style: BorderStyle::None, color: None });
    }

    #[test]
    fn test_border_style_hashable() {
        use std::collections::HashSet;
        let mut set = HashSet::new();
        set.insert(BorderStyle::None);
        set.insert(BorderStyle::Thin);
        set.insert(BorderStyle::Medium);
        set.insert(BorderStyle::Thick);
        assert_eq!(set.len(), 4);
    }

    #[test]
    fn test_cell_border_hashable() {
        use std::collections::HashSet;
        let mut set = HashSet::new();
        set.insert(CellBorder::default());
        set.insert(CellBorder::thin());
        set.insert(CellBorder { style: BorderStyle::Medium, color: None });
        set.insert(CellBorder { style: BorderStyle::Thin, color: Some([255, 0, 0, 255]) });
        assert_eq!(set.len(), 4);
    }

    // ======== Phase 1A: New fields and format primitives ========

    #[test]
    fn test_cell_format_new_field_defaults() {
        let format = CellFormat::default();
        assert_eq!(format.font_size, None);
        assert_eq!(format.font_color, None);
    }

    #[test]
    fn test_cell_format_serde_backward_compat() {
        // Old .vgrid files won't have font_size, font_color, or style_id.
        // Verify deserialization fills defaults.
        let json = r#"{"bold":true,"italic":false,"underline":false,"strikethrough":false,"alignment":"General","vertical_alignment":"Middle","text_overflow":"Clip","number_format":"General"}"#;
        let format: CellFormat = serde_json::from_str(json).unwrap();
        assert!(format.bold);
        assert_eq!(format.font_size, None);
        assert_eq!(format.font_color, None);
    }

    #[test]
    fn test_cell_style_id_serde_backward_compat() {
        let json = r#"{"value":{"Empty":null},"format":{"bold":false,"italic":false,"underline":false,"strikethrough":false,"alignment":"General","vertical_alignment":"Middle","text_overflow":"Clip","number_format":"General"}}"#;
        let cell: Cell = serde_json::from_str(json).unwrap();
        assert_eq!(cell.style_id, None);
    }

    #[test]
    fn test_cell_style_id_roundtrip() {
        let mut cell = Cell::default();
        cell.style_id = Some(42);
        let json = serde_json::to_string(&cell).unwrap();
        let restored: Cell = serde_json::from_str(&json).unwrap();
        assert_eq!(restored.style_id, Some(42));
    }

    #[test]
    fn test_merge_override_empty() {
        let base = CellFormat {
            bold: true,
            font_size: Some(14.0),
            font_color: Some([255, 0, 0, 255]),
            ..Default::default()
        };
        let ovr = CellFormatOverride::default();
        let merged = base.merge_override(&ovr);
        assert_eq!(merged.bold, true);
        assert_eq!(merged.font_size, Some(14.0));
        assert_eq!(merged.font_color, Some([255, 0, 0, 255]));
    }

    #[test]
    fn test_merge_override_replaces_fields() {
        let base = CellFormat {
            bold: true,
            font_size: Some(14.0),
            ..Default::default()
        };
        let ovr = CellFormatOverride {
            bold: Some(false),
            font_size: Some(Some(18.0)),
            font_color: Some(Some([0, 0, 255, 255])),
            ..Default::default()
        };
        let merged = base.merge_override(&ovr);
        assert_eq!(merged.bold, false);
        assert_eq!(merged.font_size, Some(18.0));
        assert_eq!(merged.font_color, Some([0, 0, 255, 255]));
    }

    #[test]
    fn test_merge_override_clears_option_fields() {
        let base = CellFormat {
            font_size: Some(14.0),
            font_color: Some([255, 0, 0, 255]),
            ..Default::default()
        };
        // Some(None) means "explicitly reset to inherit"
        let ovr = CellFormatOverride {
            font_size: Some(None),
            font_color: Some(None),
            ..Default::default()
        };
        let merged = base.merge_override(&ovr);
        assert_eq!(merged.font_size, None);
        assert_eq!(merged.font_color, None);
    }

    #[test]
    fn test_from_format_roundtrip() {
        let format = CellFormat {
            bold: true,
            italic: true,
            font_size: Some(16.0),
            font_color: Some([128, 0, 255, 255]),
            number_format: NumberFormat::currency(2),
            ..Default::default()
        };
        let ovr = CellFormatOverride::from_format(&format);
        let base = CellFormat::default();
        let merged = base.merge_override(&ovr);
        assert_eq!(merged, format);
    }

    #[test]
    fn test_has_formatting_font_size() {
        let mut format = CellFormat::default();
        assert!(!format.bold && format.font_size.is_none());
        format.font_size = Some(16.0);
        assert!(format.font_size.is_some());
    }

    #[test]
    fn test_has_formatting_font_color() {
        let mut format = CellFormat::default();
        assert!(format.font_color.is_none());
        format.font_color = Some([255, 0, 0, 255]);
        assert!(format.font_color.is_some());
    }

    #[test]
    fn test_center_across_selection_alignment() {
        let format = CellFormat {
            alignment: Alignment::CenterAcrossSelection,
            ..Default::default()
        };
        assert_eq!(format.alignment, Alignment::CenterAcrossSelection);
        assert_ne!(format.alignment, Alignment::Center);
    }

    #[test]
    fn test_number_format_custom() {
        let nf = NumberFormat::Custom("#,##0.00".to_string());
        assert_eq!(nf, NumberFormat::Custom("#,##0.00".to_string()));
        assert_ne!(nf, NumberFormat::General);
    }

    // ======== Custom number format tests ========

    #[test]
    fn test_custom_format_thousands_no_decimals() {
        assert_eq!(format_with_custom_code(1234567.0, "#,##0"), "1,234,567");
    }

    #[test]
    fn test_custom_format_thousands_two_decimals() {
        assert_eq!(format_with_custom_code(1234567.89, "#,##0.00"), "1,234,567.89");
    }

    #[test]
    fn test_custom_format_currency() {
        assert_eq!(format_with_custom_code(5000.0, "$#,##0"), "$5,000");
    }

    #[test]
    fn test_custom_format_currency_decimals() {
        assert_eq!(format_with_custom_code(1234.5, "$#,##0.00"), "$1,234.50");
    }

    #[test]
    fn test_custom_format_negative_parentheses() {
        assert_eq!(format_with_custom_code(-1234.0, "(#,##0.00)"), "(1,234.00)");
    }

    #[test]
    fn test_custom_format_negative_parentheses_positive() {
        // Positive number should not get parens
        assert_eq!(format_with_custom_code(1234.0, "(#,##0.00)"), "1,234.00");
    }

    #[test]
    fn test_custom_format_percent() {
        assert_eq!(format_with_custom_code(0.15, "0%"), "15%");
    }

    #[test]
    fn test_custom_format_percent_two_decimals() {
        assert_eq!(format_with_custom_code(0.1567, "0.00%"), "15.67%");
    }

    #[test]
    fn test_custom_format_no_thousands() {
        assert_eq!(format_with_custom_code(1234.56, "0.00"), "1234.56");
    }

    #[test]
    fn test_custom_format_zero() {
        assert_eq!(format_with_custom_code(0.0, "#,##0.00"), "0.00");
    }

    #[test]
    fn test_custom_format_negative_currency() {
        assert_eq!(format_with_custom_code(-500.0, "$#,##0.00"), "-$500.00");
    }

    #[test]
    fn test_custom_format_multi_section() {
        // positive;negative;zero format
        let code = "#,##0.00;(#,##0.00);-";
        assert_eq!(format_with_custom_code(1234.5, code), "1,234.50");
        assert_eq!(format_with_custom_code(-1234.5, code), "(1,234.50)");
        assert_eq!(format_with_custom_code(0.0, code), "-");
    }

    #[test]
    fn test_custom_format_quoted_dollar() {
        // Excel stores "$"#,##0 — dollar in quotes
        let code = r##""$"#,##0"##;
        assert_eq!(format_with_custom_code(44100.0, code), "$44,100");
        assert_eq!(format_with_custom_code(1235160.0, code), "$1,235,160");
    }

    #[test]
    fn test_custom_format_quoted_dollar_decimals() {
        let code = r##""$"#,##0.00"##;
        assert_eq!(format_with_custom_code(1234.5, code), "$1,234.50");
    }

    #[test]
    fn test_custom_format_quoted_negative_dollar() {
        // Negative with quoted dollar in parens: ("$"#,##0.00)
        let code = r##""$"#,##0.00;("$"#,##0.00)"##;
        assert_eq!(format_with_custom_code(1234.5, code), "$1,234.50");
        assert_eq!(format_with_custom_code(-1234.5, code), "($1,234.50)");
    }

    #[test]
    fn test_custom_format_quoted_suffix() {
        // Number with quoted suffix: #,##0" USD"
        let code = r##"#,##0" USD""##;
        assert_eq!(format_with_custom_code(5000.0, code), "5,000 USD");
    }

    #[test]
    fn test_custom_format_backslash_escape() {
        // Backslash-escaped literal: \$#,##0
        assert_eq!(format_with_custom_code(5000.0, r#"\$#,##0"#), "$5,000");
    }

    #[test]
    fn test_custom_format_accounting_padding() {
        // Accounting format with _ padding: _("$"* #,##0.00_)
        // The _ and * chars should be stripped, leaving ($#,##0.00)
        let code = r##"_("$"* #,##0.00_)"##;
        assert_eq!(format_with_custom_code(1234.5, code), "$1,234.50");
    }

    // --- ssfmt integration tests (via format_number → NumberFormat::Custom) ---

    #[test]
    fn test_ssfmt_accounting_zero_shows_dash() {
        // Accounting format: zero section should show "-" not "$-??"
        let code = r##"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"##;
        let result = CellValue::format_number(0.0, &NumberFormat::Custom(code.to_string()));
        // ssfmt should select the zero section and render a dash
        assert!(!result.contains("??"), "zero should not show raw ?? placeholders: got {}", result);
        assert!(result.contains("-"), "zero section should contain a dash: got {}", result);
    }

    #[test]
    fn test_ssfmt_accounting_positive() {
        let code = r##"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"##;
        let result = CellValue::format_number(1234.5, &NumberFormat::Custom(code.to_string()));
        assert!(result.contains("1,234.50"), "positive value should be formatted: got {}", result);
    }

    #[test]
    fn test_ssfmt_percent_basic() {
        let result = CellValue::format_number(0.1, &NumberFormat::Custom("0%".to_string()));
        assert_eq!(result, "10%");
    }

    #[test]
    fn test_ssfmt_thousands_decimals() {
        let result = CellValue::format_number(1234567.89, &NumberFormat::Custom("#,##0.00".to_string()));
        assert_eq!(result, "1,234,567.89");
    }

    #[test]
    fn test_strip_format_quotes() {
        assert_eq!(strip_format_quotes(r##""$"#,##0"##), "$#,##0");
        assert_eq!(strip_format_quotes(r##"#,##0" USD""##), "#,##0 USD");
        assert_eq!(strip_format_quotes(r##""$"#,##0.00;("$"#,##0.00)"##), "$#,##0.00;($#,##0.00)");
        assert_eq!(strip_format_quotes(r#"\$#,##0"#), "$#,##0");
        assert_eq!(strip_format_quotes("#,##0"), "#,##0"); // no quotes, unchanged
    }

    #[test]
    fn test_split_format_parts() {
        assert_eq!(split_format_parts("$#,##0.00"), ("$", "#,##0.00", ""));
        assert_eq!(split_format_parts("#,##0 USD"), ("", "#,##0", " USD"));
        assert_eq!(split_format_parts("$#,##0.00 USD"), ("$", "#,##0.00", " USD"));
        assert_eq!(split_format_parts("#,##0"), ("", "#,##0", ""));
        // No format chars → treat as literal
        let (p, pat, s) = split_format_parts("-");
        assert_eq!(p, "-");
        assert_eq!(pat, "");
        assert_eq!(s, "");
    }

    // ======== NumberFormat thousands/negative/symbol tests ========

    #[test]
    fn test_number_format_thousands_separator() {
        let fmt = NumberFormat::Number { decimals: 2, thousands: true, negative: NegativeStyle::Minus };
        assert_eq!(CellValue::format_number(1234567.89, &fmt), "1,234,567.89");
    }

    #[test]
    fn test_number_format_negative_styles() {
        let base = |neg: NegativeStyle| NumberFormat::Number { decimals: 2, thousands: true, negative: neg };
        assert_eq!(CellValue::format_number(-1234.56, &base(NegativeStyle::Minus)), "-1,234.56");
        assert_eq!(CellValue::format_number(-1234.56, &base(NegativeStyle::Parens)), "(1,234.56)");
        assert_eq!(CellValue::format_number(-1234.56, &base(NegativeStyle::RedMinus)), "-1,234.56");
        assert_eq!(CellValue::format_number(-1234.56, &base(NegativeStyle::RedParens)), "(1,234.56)");
    }

    #[test]
    fn test_currency_format_custom_symbol() {
        let fmt = NumberFormat::Currency {
            decimals: 2, thousands: true, negative: NegativeStyle::Minus,
            symbol: Some("EUR ".to_string()),
        };
        assert_eq!(CellValue::format_number(1234.50, &fmt), "EUR 1,234.50");
    }

    #[test]
    fn test_currency_default_symbol() {
        let fmt = NumberFormat::Currency {
            decimals: 2, thousands: true, negative: NegativeStyle::Minus,
            symbol: None,
        };
        assert_eq!(CellValue::format_number(1234.50, &fmt), "$1,234.50");
    }

    #[test]
    fn test_number_no_thousands() {
        let fmt = NumberFormat::Number { decimals: 2, thousands: false, negative: NegativeStyle::Minus };
        assert_eq!(CellValue::format_number(1234567.89, &fmt), "1234567.89");
    }

    #[test]
    fn test_should_render_red() {
        let fmt_red_minus = NumberFormat::Number { decimals: 2, thousands: true, negative: NegativeStyle::RedMinus };
        let fmt_red_parens = NumberFormat::Number { decimals: 2, thousands: true, negative: NegativeStyle::RedParens };
        let fmt_minus = NumberFormat::Number { decimals: 2, thousands: true, negative: NegativeStyle::Minus };

        assert!(fmt_red_minus.should_render_red(-1.0));
        assert!(fmt_red_parens.should_render_red(-1.0));
        assert!(!fmt_minus.should_render_red(-1.0));
        assert!(!fmt_red_minus.should_render_red(1.0)); // positive
        assert!(!fmt_red_minus.should_render_red(0.0)); // zero
    }

    #[test]
    fn test_backward_compat_number_defaults() {
        // Old behavior: Number { decimals: 2 } with defaults = no thousands, minus
        let fmt = NumberFormat::number_compat(2);
        assert_eq!(CellValue::format_number(1234.56, &fmt), "1234.56");
        assert_eq!(CellValue::format_number(-1234.56, &fmt), "-1234.56");
    }

    #[test]
    fn test_backward_compat_currency_defaults() {
        let fmt = NumberFormat::currency_compat(2);
        assert_eq!(CellValue::format_number(1234.56, &fmt), "$1234.56");
        assert_eq!(CellValue::format_number(-1234.56, &fmt), "-$1234.56");
    }

    #[test]
    fn test_format_grouped_edge_cases() {
        assert_eq!(format_grouped(0.0, 2, true), "0.00");
        assert_eq!(format_grouped(999.0, 0, true), "999");
        assert_eq!(format_grouped(1000.0, 0, true), "1,000");
        assert_eq!(format_grouped(1000.01, 2, true), "1,000.01");
    }

    #[test]
    fn test_negative_style_int_roundtrip() {
        for i in 0..4 {
            assert_eq!(NegativeStyle::from_int(i).to_int(), i);
        }
        // Unknown values default to Minus (0)
        assert_eq!(NegativeStyle::from_int(99).to_int(), 0);
    }

    #[test]
    fn test_number_format_decimals_helper() {
        assert_eq!(NumberFormat::number(2).decimals(), Some(2));
        assert_eq!(NumberFormat::currency(0).decimals(), Some(0));
        assert_eq!((NumberFormat::Percent { decimals: 3 }).decimals(), Some(3));
        assert_eq!(NumberFormat::General.decimals(), None);
    }

    #[test]
    fn test_currency_negative_parens() {
        let fmt = NumberFormat::Currency {
            decimals: 2, thousands: true, negative: NegativeStyle::Parens,
            symbol: None,
        };
        assert_eq!(CellValue::format_number(-1234.56, &fmt), "($1,234.56)");
        assert_eq!(CellValue::format_number(1234.56, &fmt), "$1,234.56");
    }
}
