//! Validation dialog state. Extracted from app.rs 2026-07-29 (pure move).

use gpui::*;
use visigrid_engine::cell::{NumberFormat, NegativeStyle};

// ============================================================================
// Validation Dialog State (Phase 4)
// ============================================================================

/// Validation type options for the dialog dropdown
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationTypeOption {
    #[default]
    AnyValue,
    List,
    WholeNumber,
    Decimal,
}

impl ValidationTypeOption {
    pub fn label(&self) -> &'static str {
        match self {
            Self::AnyValue => "Any value",
            Self::List => "List",
            Self::WholeNumber => "Whole number",
            Self::Decimal => "Decimal",
        }
    }

    pub const ALL: &'static [ValidationTypeOption] = &[
        Self::AnyValue,
        Self::List,
        Self::WholeNumber,
        Self::Decimal,
    ];
}

/// Numeric comparison operator options
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum NumericOperatorOption {
    #[default]
    Between,
    NotBetween,
    EqualTo,
    NotEqualTo,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
}

impl NumericOperatorOption {
    pub fn label(&self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::NotBetween => "not between",
            Self::EqualTo => "equal to",
            Self::NotEqualTo => "not equal to",
            Self::GreaterThan => "greater than",
            Self::LessThan => "less than",
            Self::GreaterThanOrEqual => "greater than or equal to",
            Self::LessThanOrEqual => "less than or equal to",
        }
    }

    pub const ALL: &'static [NumericOperatorOption] = &[
        Self::Between,
        Self::NotBetween,
        Self::EqualTo,
        Self::NotEqualTo,
        Self::GreaterThan,
        Self::LessThan,
        Self::GreaterThanOrEqual,
        Self::LessThanOrEqual,
    ];

    /// Whether this operator requires two values (min/max)
    pub fn needs_two_values(&self) -> bool {
        matches!(self, Self::Between | Self::NotBetween)
    }
}

/// Paste Special type selection
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PasteType {
    #[default]
    All,      // Normal paste (values + formulas)
    Values,   // Computed values only
    Formulas, // Raw formulas with reference adjustment
    Formats,  // Cell formatting only
}

impl PasteType {
    /// All available paste types in display order
    pub fn all() -> &'static [PasteType] {
        &[PasteType::All, PasteType::Values, PasteType::Formulas, PasteType::Formats]
    }

    /// Display name for UI
    pub fn label(&self) -> &'static str {
        match self {
            PasteType::All => "All",
            PasteType::Values => "Values",
            PasteType::Formulas => "Formulas",
            PasteType::Formats => "Formats",
        }
    }

    /// Keyboard accelerator for this paste type
    pub fn accelerator(&self) -> char {
        match self {
            PasteType::All => 'A',
            PasteType::Values => 'V',
            PasteType::Formulas => 'F',
            PasteType::Formats => 'O', // fOrmats (Excel convention)
        }
    }

    /// Description for UI
    pub fn description(&self) -> &'static str {
        match self {
            PasteType::All => "Paste everything (formulas, values, and formats)",
            PasteType::Values => "Paste computed values only (no formulas)",
            PasteType::Formulas => "Paste formulas with reference adjustment",
            PasteType::Formats => "Paste cell formatting only (no values)",
        }
    }
}

/// State for the Paste Special dialog
#[derive(Debug, Clone, Default)]
pub struct PasteSpecialDialogState {
    /// Currently selected paste type
    pub selected: PasteType,
}

/// Format type selection in the number format editor
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum NumberFormatEditorType {
    #[default]
    General,
    Number,
    Currency,
    Percent,
    Date,
}

/// Per-type settings cache for the number format editor
#[derive(Clone)]
pub struct TypeSettings {
    pub decimals: u8,
    pub thousands: bool,
    pub negative: NegativeStyle,
    pub currency_symbol: String,
}

impl Default for TypeSettings {
    fn default() -> Self {
        Self {
            decimals: 2,
            thousands: true,
            negative: NegativeStyle::Minus,
            currency_symbol: String::new(),
        }
    }
}

/// State for the Number Format Editor dialog (Ctrl+1 escalation)
pub struct NumberFormatEditorState {
    pub format_type: NumberFormatEditorType,
    pub preview_value: f64,
    // Active settings (mirrors the cache entry for current type)
    pub decimals: u8,
    pub thousands: bool,
    pub negative: NegativeStyle,
    pub currency_symbol: String,
    // Per-type caches
    number_cache: TypeSettings,
    currency_cache: TypeSettings,
    percent_cache: TypeSettings,
}

impl Default for NumberFormatEditorState {
    fn default() -> Self {
        Self {
            format_type: NumberFormatEditorType::Number,
            preview_value: 1234.5678,
            decimals: 2,
            thousands: true,
            negative: NegativeStyle::Minus,
            currency_symbol: String::new(),
            number_cache: TypeSettings {
                decimals: 2,
                thousands: true,
                negative: NegativeStyle::Minus,
                currency_symbol: String::new(),
            },
            currency_cache: TypeSettings {
                decimals: 2,
                thousands: true,
                negative: NegativeStyle::Parens,
                currency_symbol: String::new(),
            },
            percent_cache: TypeSettings {
                decimals: 2,
                thousands: false,
                negative: NegativeStyle::Minus,
                currency_symbol: String::new(),
            },
        }
    }
}

impl NumberFormatEditorState {
    /// Initialize from an existing NumberFormat and a sample value
    pub fn from_number_format(fmt: &NumberFormat, sample: f64) -> Self {
        let mut state = Self::default();
        state.preview_value = sample;
        match fmt {
            NumberFormat::Number { decimals, thousands, negative } => {
                state.format_type = NumberFormatEditorType::Number;
                state.decimals = *decimals;
                state.thousands = *thousands;
                state.negative = *negative;
                state.number_cache = TypeSettings {
                    decimals: *decimals,
                    thousands: *thousands,
                    negative: *negative,
                    currency_symbol: String::new(),
                };
            }
            NumberFormat::Currency { decimals, thousands, negative, symbol } => {
                state.format_type = NumberFormatEditorType::Currency;
                state.decimals = *decimals;
                state.thousands = *thousands;
                state.negative = *negative;
                state.currency_symbol = symbol.as_deref().unwrap_or("").to_string();
                state.currency_cache = TypeSettings {
                    decimals: *decimals,
                    thousands: *thousands,
                    negative: *negative,
                    currency_symbol: symbol.as_deref().unwrap_or("").to_string(),
                };
            }
            NumberFormat::Percent { decimals } => {
                state.format_type = NumberFormatEditorType::Percent;
                state.decimals = *decimals;
                state.thousands = false;
                state.negative = NegativeStyle::Minus;
                state.percent_cache = TypeSettings {
                    decimals: *decimals,
                    thousands: false,
                    negative: NegativeStyle::Minus,
                    currency_symbol: String::new(),
                };
            }
            NumberFormat::Date { .. } => {
                state.format_type = NumberFormatEditorType::Date;
            }
            _ => {
                state.format_type = NumberFormatEditorType::General;
            }
        }
        state
    }

    /// Convert current state to a NumberFormat
    pub fn to_number_format(&self) -> NumberFormat {
        match self.format_type {
            NumberFormatEditorType::General => NumberFormat::General,
            NumberFormatEditorType::Number => NumberFormat::Number {
                decimals: self.decimals.min(10),
                thousands: self.thousands,
                negative: self.negative,
            },
            NumberFormatEditorType::Currency => NumberFormat::Currency {
                decimals: self.decimals.min(10),
                thousands: self.thousands,
                negative: self.negative,
                symbol: if self.currency_symbol.is_empty() { None } else { Some(self.currency_symbol.clone()) },
            },
            NumberFormatEditorType::Percent => NumberFormat::Percent {
                decimals: self.decimals.min(10),
            },
            NumberFormatEditorType::Date => NumberFormat::Date {
                style: visigrid_engine::cell::DateStyle::Short,
            },
        }
    }

    /// Format a value using current settings for preview
    pub fn preview(&self) -> String {
        use visigrid_engine::cell::CellValue;
        CellValue::format_number(self.preview_value, &self.to_number_format())
    }

    /// Format the negative version for preview
    pub fn preview_negative(&self) -> String {
        use visigrid_engine::cell::CellValue;
        CellValue::format_number(-self.preview_value.abs(), &self.to_number_format())
    }

    /// Format zero for preview
    pub fn preview_zero(&self) -> String {
        use visigrid_engine::cell::CellValue;
        CellValue::format_number(0.0, &self.to_number_format())
    }

    /// Switch to a different format type, preserving per-type caches
    pub fn switch_type(&mut self, new_type: NumberFormatEditorType) {
        if self.format_type == new_type {
            return;
        }
        // Save current settings to outgoing cache
        let current = TypeSettings {
            decimals: self.decimals,
            thousands: self.thousands,
            negative: self.negative,
            currency_symbol: self.currency_symbol.clone(),
        };
        match self.format_type {
            NumberFormatEditorType::Number => self.number_cache = current,
            NumberFormatEditorType::Currency => self.currency_cache = current,
            NumberFormatEditorType::Percent => self.percent_cache = current,
            _ => {}
        }
        // Restore from incoming cache
        self.format_type = new_type;
        match new_type {
            NumberFormatEditorType::Number => {
                let c = &self.number_cache;
                self.decimals = c.decimals;
                self.thousands = c.thousands;
                self.negative = c.negative;
                self.currency_symbol = c.currency_symbol.clone();
            }
            NumberFormatEditorType::Currency => {
                let c = &self.currency_cache;
                self.decimals = c.decimals;
                self.thousands = c.thousands;
                self.negative = c.negative;
                self.currency_symbol = c.currency_symbol.clone();
            }
            NumberFormatEditorType::Percent => {
                let c = &self.percent_cache;
                self.decimals = c.decimals;
                self.thousands = c.thousands;
                self.negative = c.negative;
                self.currency_symbol = c.currency_symbol.clone();
            }
            _ => {
                self.decimals = 2;
                self.thousands = false;
                self.negative = NegativeStyle::Minus;
                self.currency_symbol = String::new();
            }
        }
    }
}

/// Which field in the validation dialog has focus
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationDialogFocus {
    #[default]
    None,
    TypeDropdown,
    OperatorDropdown,
    Source,      // List source field
    Value1,      // First value (or Minimum for Between)
    Value2,      // Second value (Maximum for Between)
}

/// State for the data validation dialog (Phase 4)
#[derive(Debug, Clone, Default)]
pub struct ValidationDialogState {
    /// Currently selected validation type
    pub validation_type: ValidationTypeOption,
    /// Whether the type dropdown is expanded
    pub type_dropdown_open: bool,
    /// Whether the operator dropdown is expanded
    pub operator_dropdown_open: bool,

    // List validation fields
    /// Source for list validation (e.g., "A1:A10" or "Yes,No,Maybe")
    pub list_source: String,
    /// Show dropdown arrow in cell
    pub show_dropdown: bool,

    // Numeric validation fields
    /// Comparison operator
    pub numeric_operator: NumericOperatorOption,
    /// First value (or minimum for between)
    pub value1: String,
    /// Second value (maximum for between)
    pub value2: String,

    // Common fields
    /// Allow blank cells
    pub ignore_blank: bool,

    /// Which field currently has focus
    pub focus: ValidationDialogFocus,

    /// Error message to display (validation errors)
    pub error: Option<String>,

    /// The range we're applying validation to (captured when dialog opens)
    pub target_range: Option<visigrid_engine::validation::CellRange>,

    /// Whether we loaded existing validation (for Clear button visibility)
    pub has_existing_validation: bool,
}

impl ValidationDialogState {
    /// Reset to defaults for a new dialog session
    pub fn reset(&mut self) {
        *self = Self::default();
        self.show_dropdown = true;  // Default to showing dropdown for list
        self.ignore_blank = true;   // Default to allowing blank
    }

    /// Load state from an existing validation rule
    pub fn load_from_rule(&mut self, rule: &visigrid_engine::validation::ValidationRule) {
        use visigrid_engine::validation::{ValidationType, ListSource};

        self.reset();
        self.has_existing_validation = true;
        self.ignore_blank = rule.ignore_blank;
        self.show_dropdown = rule.show_dropdown;

        match &rule.rule_type {
            // NOTE: No AnyValue case - that variant no longer exists in engine
            ValidationType::List(source) => {
                self.validation_type = ValidationTypeOption::List;
                match source {
                    ListSource::Inline(items) => {
                        self.list_source = items.join(",");
                    }
                    ListSource::Range(r) => {
                        self.list_source = r.clone();
                    }
                    ListSource::NamedRange(n) => {
                        self.list_source = n.clone();
                    }
                }
            }
            ValidationType::WholeNumber(constraint) => {
                self.validation_type = ValidationTypeOption::WholeNumber;
                self.load_numeric_constraint(constraint);
            }
            ValidationType::Decimal(constraint) => {
                self.validation_type = ValidationTypeOption::Decimal;
                self.load_numeric_constraint(constraint);
            }
            _ => {
                // Date, Time, TextLength, Custom - not yet supported in dialog
                // Show as AnyValue (read-only)
                self.validation_type = ValidationTypeOption::AnyValue;
            }
        }
    }

    fn load_numeric_constraint(&mut self, constraint: &visigrid_engine::validation::NumericConstraint) {
        use visigrid_engine::validation::{ComparisonOperator, ConstraintValue};

        self.numeric_operator = match constraint.operator {
            ComparisonOperator::Between => NumericOperatorOption::Between,
            ComparisonOperator::NotBetween => NumericOperatorOption::NotBetween,
            ComparisonOperator::EqualTo => NumericOperatorOption::EqualTo,
            ComparisonOperator::NotEqualTo => NumericOperatorOption::NotEqualTo,
            ComparisonOperator::GreaterThan => NumericOperatorOption::GreaterThan,
            ComparisonOperator::LessThan => NumericOperatorOption::LessThan,
            ComparisonOperator::GreaterThanOrEqual => NumericOperatorOption::GreaterThanOrEqual,
            ComparisonOperator::LessThanOrEqual => NumericOperatorOption::LessThanOrEqual,
        };

        // Helper to convert constraint value to string
        let value_to_string = |v: &ConstraintValue| -> String {
            match v {
                ConstraintValue::Number(n) => n.to_string(),
                ConstraintValue::CellRef(r) => r.clone(),
                ConstraintValue::Formula(f) => f.clone(),
            }
        };

        self.value1 = value_to_string(&constraint.value1);
        if let Some(ref v2) = constraint.value2 {
            self.value2 = value_to_string(v2);
        }
    }
}

