use serde::{Deserialize, Serialize};

use super::formula::parser::{self, Expr};

/// Text alignment
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize, PartialEq)]
pub enum Alignment {
    #[default]
    Left,
    Center,
    Right,
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
    pub alignment: Alignment,
    pub number_format: NumberFormat,
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

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Cell {
    pub value: CellValue,
    pub format: CellFormat,
}

impl Cell {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn set(&mut self, input: &str) {
        self.value = CellValue::from_input(input);
    }
}
