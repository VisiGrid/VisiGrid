//! Transient numeric editor used by Preferences.
use crate::settings::CellSizeDefaults;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellSizeField {
    ColumnWidth,
    RowHeight,
    FontSize,
}

impl CellSizeField {
    pub fn value(self, sizes: CellSizeDefaults, font_size: f32) -> f32 {
        match self {
            Self::ColumnWidth => sizes.column_width,
            Self::RowHeight => sizes.row_height,
            Self::FontSize => font_size,
        }
    }
    pub fn unit(self) -> &'static str {
        if self == Self::FontSize {
            "pt"
        } else {
            "px"
        }
    }
    pub fn next(self) -> Self {
        match self {
            Self::ColumnWidth => Self::RowHeight,
            Self::RowHeight => Self::FontSize,
            Self::FontSize => Self::ColumnWidth,
        }
    }
    pub fn parse(self, text: &str) -> Result<f32, String> {
        let (min, max) = match self {
            Self::ColumnWidth => CellSizeDefaults::COLUMN_LIMITS,
            Self::RowHeight => CellSizeDefaults::ROW_LIMITS,
            Self::FontSize => crate::settings::FONT_SIZE_LIMITS,
        };
        text.trim()
            .parse::<f32>()
            .ok()
            .filter(|n| n.is_finite() && *n >= min && *n <= max)
            .ok_or_else(|| format!("Enter a size from {min} to {max} {}.", self.unit()))
    }
}

#[derive(Default)]
pub struct CellSizeInput {
    pub field: Option<CellSizeField>,
    pub text: String,
    pub all_selected: bool,
    pub error: Option<String>,
}

impl CellSizeInput {
    pub fn begin(&mut self, field: CellSizeField, sizes: CellSizeDefaults, font_size: f32) {
        self.field = Some(field);
        self.text = field.value(sizes, font_size).to_string();
        self.all_selected = true;
        self.error = None;
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn input_accepts_pixels_and_rejects_invalid_or_unsafe_sizes() {
        assert_eq!(CellSizeField::ColumnWidth.parse(" 80.5 "), Ok(80.5));
        assert_eq!(CellSizeField::RowHeight.parse("24"), Ok(24.0));
        for text in ["", "abc", "NaN", "inf", "-1", "0", "501"] {
            assert!(CellSizeField::ColumnWidth.parse(text).is_err());
        }
        assert!(CellSizeField::RowHeight.parse("11").is_err());
        assert!(CellSizeField::RowHeight.parse("201").is_err());
    }
}
