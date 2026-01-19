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

/// Number format type
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum NumberFormat {
    #[default]
    General,
    Number { decimals: u8 },
    Currency { decimals: u8 },
    Percent { decimals: u8 },
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
