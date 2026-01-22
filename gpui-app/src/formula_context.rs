//! Formula Context Analyzer
//!
//! This module provides the single source of truth for formula UI features.
//! All autocomplete, signature help, F4 cycling, and error highlighting
//! route through the `analyze()` function.
//!
//! See docs/gpui/formulas.md for the full specification.

use std::ops::Range;

// ============================================================================
// Core Types
// ============================================================================

/// Describes the editing context at a specific cursor position within a formula.
#[derive(Debug, Clone)]
pub struct FormulaContext {
    /// What kind of position the cursor is in
    pub mode: FormulaEditMode,

    /// Cursor position (char index from start of formula, including '=')
    pub cursor: usize,

    /// If inside a function's argument list, which function
    pub current_function: Option<&'static FunctionInfo>,

    /// If inside argument list, which argument (0-indexed)
    pub current_arg_index: Option<usize>,

    /// The token spanning the cursor position, if any
    pub token_at_cursor: Option<TokenSpan>,

    /// The primary span for operations (F4 cycles this, hover attaches here)
    pub primary_span: Option<Range<usize>>,

    /// What range autocomplete should replace when accepting a suggestion
    pub replace_range: Range<usize>,

    /// Nesting depth of parentheses at cursor
    pub paren_depth: usize,

    /// The identifier text at cursor (for autocomplete filtering)
    pub identifier_text: Option<String>,
}

/// The mode determines what UI behavior is appropriate
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaEditMode {
    /// Right after '=' with nothing typed
    Start,
    /// Typing a function name (or could be cell ref)
    Identifier,
    /// Inside a function's argument list (after '(' or ',')
    ArgList,
    /// Inside a string literal
    String,
    /// On a cell reference or range
    Reference,
    /// On an operator (+, -, *, /, &, <, >, =)
    Operator,
    /// On a number literal
    Number,
    /// After a complete expression (e.g., after closing paren)
    Complete,
}

/// A token with its position in the formula
#[derive(Debug, Clone)]
pub struct TokenSpan {
    pub token_type: TokenType,
    pub range: Range<usize>,  // Char indices
    pub text: String,
}

/// Token types for syntax highlighting and context detection
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TokenType {
    Function,
    CellRef,
    Range,
    NamedRange,
    Number,
    String,
    Boolean,
    Operator,
    Comparison,
    Paren,
    Comma,
    Colon,
    Error,
    // Structural tokens
    Whitespace,
    Bang,           // '!' for sheet references
    Percent,        // '%' suffix
    UnaryMinus,     // Leading minus (distinguished from subtraction)
}

// ============================================================================
// Function Metadata
// ============================================================================

/// Information about a formula function
#[derive(Debug, Clone)]
pub struct FunctionInfo {
    pub name: &'static str,
    pub signature: &'static str,
    pub description: &'static str,
    pub category: FunctionCategory,
    pub parameters: &'static [ParameterInfo],
}

/// Parameter information for signature help
#[derive(Debug, Clone)]
pub struct ParameterInfo {
    pub name: &'static str,
    pub description: &'static str,
    pub optional: bool,
    pub repeatable: bool,
}

/// Function categories
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FunctionCategory {
    Math,
    Logical,
    Text,
    Lookup,
    DateTime,
    Statistical,
    Array,
    Conditional,
    Trigonometry,
}

// ============================================================================
// Function Database (96 functions)
// ============================================================================

/// All supported functions with their metadata
pub static FUNCTIONS: &[FunctionInfo] = &[
    // Math (23)
    FunctionInfo {
        name: "SUM",
        signature: "SUM(number1, [number2], ...)",
        description: "Adds all the numbers in a range of cells.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range to add.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges to add.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "AVERAGE",
        signature: "AVERAGE(number1, [number2], ...)",
        description: "Returns the average of the arguments.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "MIN",
        signature: "MIN(number1, [number2], ...)",
        description: "Returns the smallest value in a set of values.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "MAX",
        signature: "MAX(number1, [number2], ...)",
        description: "Returns the largest value in a set of values.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "COUNT",
        signature: "COUNT(value1, [value2], ...)",
        description: "Counts the number of cells that contain numbers.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "value1", description: "The first value or range.", optional: false, repeatable: false },
            ParameterInfo { name: "value2", description: "Additional values or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "COUNTA",
        signature: "COUNTA(value1, [value2], ...)",
        description: "Counts the number of non-empty cells.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "value1", description: "The first value or range.", optional: false, repeatable: false },
            ParameterInfo { name: "value2", description: "Additional values or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "ABS",
        signature: "ABS(number)",
        description: "Returns the absolute value of a number.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The number to get the absolute value of.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ROUND",
        signature: "ROUND(number, num_digits)",
        description: "Rounds a number to a specified number of digits.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The number to round.", optional: false, repeatable: false },
            ParameterInfo { name: "num_digits", description: "The number of digits to round to.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "INT",
        signature: "INT(number)",
        description: "Rounds a number down to the nearest integer.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The number to round down.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "MOD",
        signature: "MOD(number, divisor)",
        description: "Returns the remainder after division.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The number to divide.", optional: false, repeatable: false },
            ParameterInfo { name: "divisor", description: "The number to divide by.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "POWER",
        signature: "POWER(number, power)",
        description: "Returns the result of a number raised to a power.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The base number.", optional: false, repeatable: false },
            ParameterInfo { name: "power", description: "The exponent.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "SQRT",
        signature: "SQRT(number)",
        description: "Returns the square root of a number.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The number to get the square root of.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "CEILING",
        signature: "CEILING(number, significance)",
        description: "Rounds a number up to the nearest multiple of significance.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The number to round.", optional: false, repeatable: false },
            ParameterInfo { name: "significance", description: "The multiple to round to.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "FLOOR",
        signature: "FLOOR(number, significance)",
        description: "Rounds a number down to the nearest multiple of significance.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The number to round.", optional: false, repeatable: false },
            ParameterInfo { name: "significance", description: "The multiple to round to.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "PRODUCT",
        signature: "PRODUCT(number1, [number2], ...)",
        description: "Multiplies all the numbers given as arguments.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "MEDIAN",
        signature: "MEDIAN(number1, [number2], ...)",
        description: "Returns the median of the given numbers.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "LOG",
        signature: "LOG(number, [base])",
        description: "Returns the logarithm of a number to a specified base.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The positive number.", optional: false, repeatable: false },
            ParameterInfo { name: "base", description: "The base (default 10).", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "LOG10",
        signature: "LOG10(number)",
        description: "Returns the base-10 logarithm of a number.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The positive number.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "LN",
        signature: "LN(number)",
        description: "Returns the natural logarithm of a number.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The positive number.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "EXP",
        signature: "EXP(number)",
        description: "Returns e raised to the power of a number.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "number", description: "The exponent.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "RAND",
        signature: "RAND()",
        description: "Returns a random number between 0 and 1.",
        category: FunctionCategory::Math,
        parameters: &[],
    },
    FunctionInfo {
        name: "RANDBETWEEN",
        signature: "RANDBETWEEN(bottom, top)",
        description: "Returns a random integer between two numbers.",
        category: FunctionCategory::Math,
        parameters: &[
            ParameterInfo { name: "bottom", description: "The smallest integer.", optional: false, repeatable: false },
            ParameterInfo { name: "top", description: "The largest integer.", optional: false, repeatable: false },
        ],
    },

    // Logical (12)
    FunctionInfo {
        name: "IF",
        signature: "IF(logical_test, value_if_true, [value_if_false])",
        description: "Returns one value if a condition is true and another if false.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "logical_test", description: "The condition to test.", optional: false, repeatable: false },
            ParameterInfo { name: "value_if_true", description: "The value if true.", optional: false, repeatable: false },
            ParameterInfo { name: "value_if_false", description: "The value if false.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "AND",
        signature: "AND(logical1, [logical2], ...)",
        description: "Returns TRUE if all arguments are TRUE.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "logical1", description: "The first condition.", optional: false, repeatable: false },
            ParameterInfo { name: "logical2", description: "Additional conditions.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "OR",
        signature: "OR(logical1, [logical2], ...)",
        description: "Returns TRUE if any argument is TRUE.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "logical1", description: "The first condition.", optional: false, repeatable: false },
            ParameterInfo { name: "logical2", description: "Additional conditions.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "NOT",
        signature: "NOT(logical)",
        description: "Reverses the logic of its argument.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "logical", description: "The value to reverse.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "IFERROR",
        signature: "IFERROR(value, value_if_error)",
        description: "Returns a value if an expression results in an error.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "value", description: "The value to check for error.", optional: false, repeatable: false },
            ParameterInfo { name: "value_if_error", description: "The value to return if error.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ISBLANK",
        signature: "ISBLANK(value)",
        description: "Returns TRUE if the value is blank.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "value", description: "The value to check.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ISNUMBER",
        signature: "ISNUMBER(value)",
        description: "Returns TRUE if the value is a number.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "value", description: "The value to check.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ISTEXT",
        signature: "ISTEXT(value)",
        description: "Returns TRUE if the value is text.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "value", description: "The value to check.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ISERROR",
        signature: "ISERROR(value)",
        description: "Returns TRUE if the value is an error.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "value", description: "The value to check.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ISNA",
        signature: "ISNA(value)",
        description: "Returns TRUE if the value is the #N/A error.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "value", description: "The value to check.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "IFS",
        signature: "IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)",
        description: "Checks multiple conditions and returns the first true result.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "logical_test1", description: "First condition.", optional: false, repeatable: false },
            ParameterInfo { name: "value_if_true1", description: "Value if first condition is true.", optional: false, repeatable: false },
            ParameterInfo { name: "logical_test2", description: "Additional conditions.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "SWITCH",
        signature: "SWITCH(expression, value1, result1, [value2, result2], ..., [default])",
        description: "Evaluates an expression against a list of values.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "expression", description: "The value to match.", optional: false, repeatable: false },
            ParameterInfo { name: "value1", description: "First value to match against.", optional: false, repeatable: false },
            ParameterInfo { name: "result1", description: "Result if first value matches.", optional: false, repeatable: false },
            ParameterInfo { name: "default", description: "Default value if no match.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "CHOOSE",
        signature: "CHOOSE(index_num, value1, [value2], ...)",
        description: "Chooses a value from a list based on an index number.",
        category: FunctionCategory::Logical,
        parameters: &[
            ParameterInfo { name: "index_num", description: "The index (1-based).", optional: false, repeatable: false },
            ParameterInfo { name: "value1", description: "The first value.", optional: false, repeatable: false },
            ParameterInfo { name: "value2", description: "Additional values.", optional: true, repeatable: true },
        ],
    },

    // Text (14)
    FunctionInfo {
        name: "CONCATENATE",
        signature: "CONCATENATE(text1, [text2], ...)",
        description: "Joins several text strings into one string.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text1", description: "The first text.", optional: false, repeatable: false },
            ParameterInfo { name: "text2", description: "Additional text.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "CONCAT",
        signature: "CONCAT(text1, [text2], ...)",
        description: "Joins text strings (modern version of CONCATENATE).",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text1", description: "The first text or range.", optional: false, repeatable: false },
            ParameterInfo { name: "text2", description: "Additional text or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "LEFT",
        signature: "LEFT(text, [num_chars])",
        description: "Returns the leftmost characters from a text string.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text string.", optional: false, repeatable: false },
            ParameterInfo { name: "num_chars", description: "Number of characters (default 1).", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "RIGHT",
        signature: "RIGHT(text, [num_chars])",
        description: "Returns the rightmost characters from a text string.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text string.", optional: false, repeatable: false },
            ParameterInfo { name: "num_chars", description: "Number of characters (default 1).", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "MID",
        signature: "MID(text, start_num, num_chars)",
        description: "Returns characters from the middle of a text string.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text string.", optional: false, repeatable: false },
            ParameterInfo { name: "start_num", description: "The starting position (1-based).", optional: false, repeatable: false },
            ParameterInfo { name: "num_chars", description: "Number of characters.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "LEN",
        signature: "LEN(text)",
        description: "Returns the number of characters in a text string.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text string.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "UPPER",
        signature: "UPPER(text)",
        description: "Converts text to uppercase.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text to convert.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "LOWER",
        signature: "LOWER(text)",
        description: "Converts text to lowercase.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text to convert.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "TRIM",
        signature: "TRIM(text)",
        description: "Removes extra spaces from text.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text to trim.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "TEXT",
        signature: "TEXT(value, format_text)",
        description: "Formats a number as text with a specified format.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "value", description: "The number to format.", optional: false, repeatable: false },
            ParameterInfo { name: "format_text", description: "The format code.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "VALUE",
        signature: "VALUE(text)",
        description: "Converts a text string to a number.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text to convert.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "FIND",
        signature: "FIND(find_text, within_text, [start_num])",
        description: "Finds one text string within another (case-sensitive).",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "find_text", description: "The text to find.", optional: false, repeatable: false },
            ParameterInfo { name: "within_text", description: "The text to search in.", optional: false, repeatable: false },
            ParameterInfo { name: "start_num", description: "Starting position (default 1).", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "SUBSTITUTE",
        signature: "SUBSTITUTE(text, old_text, new_text, [instance_num])",
        description: "Replaces old text with new text in a string.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text to modify.", optional: false, repeatable: false },
            ParameterInfo { name: "old_text", description: "The text to replace.", optional: false, repeatable: false },
            ParameterInfo { name: "new_text", description: "The replacement text.", optional: false, repeatable: false },
            ParameterInfo { name: "instance_num", description: "Which occurrence to replace.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "REPT",
        signature: "REPT(text, number_times)",
        description: "Repeats text a given number of times.",
        category: FunctionCategory::Text,
        parameters: &[
            ParameterInfo { name: "text", description: "The text to repeat.", optional: false, repeatable: false },
            ParameterInfo { name: "number_times", description: "Number of repetitions.", optional: false, repeatable: false },
        ],
    },

    // Conditional (3)
    FunctionInfo {
        name: "SUMIF",
        signature: "SUMIF(range, criteria, [sum_range])",
        description: "Sums cells that meet a criteria.",
        category: FunctionCategory::Conditional,
        parameters: &[
            ParameterInfo { name: "range", description: "The range to evaluate.", optional: false, repeatable: false },
            ParameterInfo { name: "criteria", description: "The criteria to match.", optional: false, repeatable: false },
            ParameterInfo { name: "sum_range", description: "The cells to sum.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "COUNTIF",
        signature: "COUNTIF(range, criteria)",
        description: "Counts cells that meet a criteria.",
        category: FunctionCategory::Conditional,
        parameters: &[
            ParameterInfo { name: "range", description: "The range to evaluate.", optional: false, repeatable: false },
            ParameterInfo { name: "criteria", description: "The criteria to match.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "COUNTBLANK",
        signature: "COUNTBLANK(range)",
        description: "Counts the number of blank cells in a range.",
        category: FunctionCategory::Conditional,
        parameters: &[
            ParameterInfo { name: "range", description: "The range to count.", optional: false, repeatable: false },
        ],
    },

    // Lookup (8)
    FunctionInfo {
        name: "VLOOKUP",
        signature: "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
        description: "Looks for a value in the leftmost column and returns a value in the same row.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "lookup_value", description: "The value to search for.", optional: false, repeatable: false },
            ParameterInfo { name: "table_array", description: "The range containing the data.", optional: false, repeatable: false },
            ParameterInfo { name: "col_index_num", description: "The column number to return (1-indexed).", optional: false, repeatable: false },
            ParameterInfo { name: "range_lookup", description: "TRUE for approximate, FALSE for exact.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "HLOOKUP",
        signature: "HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])",
        description: "Looks for a value in the top row and returns a value in the same column.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "lookup_value", description: "The value to search for.", optional: false, repeatable: false },
            ParameterInfo { name: "table_array", description: "The range containing the data.", optional: false, repeatable: false },
            ParameterInfo { name: "row_index_num", description: "The row number to return (1-indexed).", optional: false, repeatable: false },
            ParameterInfo { name: "range_lookup", description: "TRUE for approximate, FALSE for exact.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "INDEX",
        signature: "INDEX(array, row_num, [column_num])",
        description: "Returns the value at a given position in a range.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "array", description: "The range of cells.", optional: false, repeatable: false },
            ParameterInfo { name: "row_num", description: "The row number.", optional: false, repeatable: false },
            ParameterInfo { name: "column_num", description: "The column number.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "MATCH",
        signature: "MATCH(lookup_value, lookup_array, [match_type])",
        description: "Returns the position of a value in a range.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "lookup_value", description: "The value to find.", optional: false, repeatable: false },
            ParameterInfo { name: "lookup_array", description: "The range to search.", optional: false, repeatable: false },
            ParameterInfo { name: "match_type", description: "1 (less than), 0 (exact), -1 (greater than).", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ROW",
        signature: "ROW([reference])",
        description: "Returns the row number of a reference.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "reference", description: "The cell reference.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "COLUMN",
        signature: "COLUMN([reference])",
        description: "Returns the column number of a reference.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "reference", description: "The cell reference.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ROWS",
        signature: "ROWS(array)",
        description: "Returns the number of rows in a reference.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "array", description: "The range.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "COLUMNS",
        signature: "COLUMNS(array)",
        description: "Returns the number of columns in a reference.",
        category: FunctionCategory::Lookup,
        parameters: &[
            ParameterInfo { name: "array", description: "The range.", optional: false, repeatable: false },
        ],
    },

    // Date/Time (13)
    FunctionInfo {
        name: "TODAY",
        signature: "TODAY()",
        description: "Returns the current date.",
        category: FunctionCategory::DateTime,
        parameters: &[],
    },
    FunctionInfo {
        name: "NOW",
        signature: "NOW()",
        description: "Returns the current date and time.",
        category: FunctionCategory::DateTime,
        parameters: &[],
    },
    FunctionInfo {
        name: "DATE",
        signature: "DATE(year, month, day)",
        description: "Creates a date from year, month, and day.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "year", description: "The year.", optional: false, repeatable: false },
            ParameterInfo { name: "month", description: "The month (1-12).", optional: false, repeatable: false },
            ParameterInfo { name: "day", description: "The day (1-31).", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "YEAR",
        signature: "YEAR(serial_number)",
        description: "Returns the year from a date.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "serial_number", description: "The date.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "MONTH",
        signature: "MONTH(serial_number)",
        description: "Returns the month from a date.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "serial_number", description: "The date.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "DAY",
        signature: "DAY(serial_number)",
        description: "Returns the day from a date.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "serial_number", description: "The date.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "WEEKDAY",
        signature: "WEEKDAY(serial_number, [return_type])",
        description: "Returns the day of the week from a date.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "serial_number", description: "The date.", optional: false, repeatable: false },
            ParameterInfo { name: "return_type", description: "1=Sun-Sat, 2=Mon-Sun, 3=0-6.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "DATEDIF",
        signature: "DATEDIF(start_date, end_date, unit)",
        description: "Calculates the difference between two dates.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "start_date", description: "The start date.", optional: false, repeatable: false },
            ParameterInfo { name: "end_date", description: "The end date.", optional: false, repeatable: false },
            ParameterInfo { name: "unit", description: "Y, M, D, YM, YD, or MD.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "EDATE",
        signature: "EDATE(start_date, months)",
        description: "Returns a date a specified number of months away.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "start_date", description: "The start date.", optional: false, repeatable: false },
            ParameterInfo { name: "months", description: "Number of months to add.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "EOMONTH",
        signature: "EOMONTH(start_date, months)",
        description: "Returns the last day of a month a specified number of months away.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "start_date", description: "The start date.", optional: false, repeatable: false },
            ParameterInfo { name: "months", description: "Number of months to add.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "HOUR",
        signature: "HOUR(serial_number)",
        description: "Returns the hour from a time value.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "serial_number", description: "The time.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "MINUTE",
        signature: "MINUTE(serial_number)",
        description: "Returns the minute from a time value.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "serial_number", description: "The time.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "SECOND",
        signature: "SECOND(serial_number)",
        description: "Returns the second from a time value.",
        category: FunctionCategory::DateTime,
        parameters: &[
            ParameterInfo { name: "serial_number", description: "The time.", optional: false, repeatable: false },
        ],
    },

    // Trigonometry (10)
    FunctionInfo {
        name: "PI",
        signature: "PI()",
        description: "Returns the value of pi (3.14159...).",
        category: FunctionCategory::Trigonometry,
        parameters: &[],
    },
    FunctionInfo {
        name: "SIN",
        signature: "SIN(number)",
        description: "Returns the sine of an angle (in radians).",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "number", description: "The angle in radians.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "COS",
        signature: "COS(number)",
        description: "Returns the cosine of an angle (in radians).",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "number", description: "The angle in radians.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "TAN",
        signature: "TAN(number)",
        description: "Returns the tangent of an angle (in radians).",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "number", description: "The angle in radians.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ASIN",
        signature: "ASIN(number)",
        description: "Returns the arcsine of a number (in radians).",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "number", description: "The sine value (-1 to 1).", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ACOS",
        signature: "ACOS(number)",
        description: "Returns the arccosine of a number (in radians).",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "number", description: "The cosine value (-1 to 1).", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ATAN",
        signature: "ATAN(number)",
        description: "Returns the arctangent of a number (in radians).",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "number", description: "The tangent value.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "ATAN2",
        signature: "ATAN2(x_num, y_num)",
        description: "Returns the arctangent from x and y coordinates.",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "x_num", description: "The x coordinate.", optional: false, repeatable: false },
            ParameterInfo { name: "y_num", description: "The y coordinate.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "DEGREES",
        signature: "DEGREES(angle)",
        description: "Converts radians to degrees.",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "angle", description: "The angle in radians.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "RADIANS",
        signature: "RADIANS(angle)",
        description: "Converts degrees to radians.",
        category: FunctionCategory::Trigonometry,
        parameters: &[
            ParameterInfo { name: "angle", description: "The angle in degrees.", optional: false, repeatable: false },
        ],
    },

    // Statistical (8)
    FunctionInfo {
        name: "STDEV",
        signature: "STDEV(number1, [number2], ...)",
        description: "Estimates standard deviation based on a sample.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "STDEV.S",
        signature: "STDEV.S(number1, [number2], ...)",
        description: "Estimates standard deviation based on a sample.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "STDEV.P",
        signature: "STDEV.P(number1, [number2], ...)",
        description: "Calculates standard deviation based on an entire population.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "STDEVP",
        signature: "STDEVP(number1, [number2], ...)",
        description: "Calculates standard deviation based on an entire population.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "VAR",
        signature: "VAR(number1, [number2], ...)",
        description: "Estimates variance based on a sample.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "VAR.S",
        signature: "VAR.S(number1, [number2], ...)",
        description: "Estimates variance based on a sample.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "VAR.P",
        signature: "VAR.P(number1, [number2], ...)",
        description: "Calculates variance based on an entire population.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },
    FunctionInfo {
        name: "VARP",
        signature: "VARP(number1, [number2], ...)",
        description: "Calculates variance based on an entire population.",
        category: FunctionCategory::Statistical,
        parameters: &[
            ParameterInfo { name: "number1", description: "The first number or range.", optional: false, repeatable: false },
            ParameterInfo { name: "number2", description: "Additional numbers or ranges.", optional: true, repeatable: true },
        ],
    },

    // Array (5)
    FunctionInfo {
        name: "SEQUENCE",
        signature: "SEQUENCE(rows, [columns], [start], [step])",
        description: "Generates a sequence of numbers.",
        category: FunctionCategory::Array,
        parameters: &[
            ParameterInfo { name: "rows", description: "Number of rows.", optional: false, repeatable: false },
            ParameterInfo { name: "columns", description: "Number of columns (default 1).", optional: true, repeatable: false },
            ParameterInfo { name: "start", description: "Starting value (default 1).", optional: true, repeatable: false },
            ParameterInfo { name: "step", description: "Step value (default 1).", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "TRANSPOSE",
        signature: "TRANSPOSE(array)",
        description: "Transposes the rows and columns of an array.",
        category: FunctionCategory::Array,
        parameters: &[
            ParameterInfo { name: "array", description: "The array to transpose.", optional: false, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "SORT",
        signature: "SORT(array, [sort_index], [sort_order], [by_col])",
        description: "Sorts the contents of a range or array.",
        category: FunctionCategory::Array,
        parameters: &[
            ParameterInfo { name: "array", description: "The range to sort.", optional: false, repeatable: false },
            ParameterInfo { name: "sort_index", description: "Column/row to sort by.", optional: true, repeatable: false },
            ParameterInfo { name: "sort_order", description: "1=ascending, -1=descending.", optional: true, repeatable: false },
            ParameterInfo { name: "by_col", description: "TRUE to sort by column.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "UNIQUE",
        signature: "UNIQUE(array, [by_col], [exactly_once])",
        description: "Returns the unique values from a range.",
        category: FunctionCategory::Array,
        parameters: &[
            ParameterInfo { name: "array", description: "The range.", optional: false, repeatable: false },
            ParameterInfo { name: "by_col", description: "TRUE to compare columns.", optional: true, repeatable: false },
            ParameterInfo { name: "exactly_once", description: "TRUE for values that appear only once.", optional: true, repeatable: false },
        ],
    },
    FunctionInfo {
        name: "FILTER",
        signature: "FILTER(array, include, [if_empty])",
        description: "Filters a range based on criteria.",
        category: FunctionCategory::Array,
        parameters: &[
            ParameterInfo { name: "array", description: "The range to filter.", optional: false, repeatable: false },
            ParameterInfo { name: "include", description: "Boolean array of same height/width.", optional: false, repeatable: false },
            ParameterInfo { name: "if_empty", description: "Value if no results.", optional: true, repeatable: false },
        ],
    },
];

// ============================================================================
// Helper Functions
// ============================================================================

/// Look up a function by name (case-insensitive)
pub fn get_function(name: &str) -> Option<&'static FunctionInfo> {
    let upper = name.to_ascii_uppercase();
    FUNCTIONS.iter().find(|f| f.name == upper)
}

/// Get all functions matching a prefix (for autocomplete)
pub fn get_functions_by_prefix(prefix: &str) -> Vec<&'static FunctionInfo> {
    let upper = prefix.to_ascii_uppercase();
    FUNCTIONS.iter()
        .filter(|f| f.name.starts_with(&upper))
        .collect()
}

/// Convert char index to byte offset
pub fn char_to_byte(s: &str, char_idx: usize) -> usize {
    s.char_indices()
        .nth(char_idx)
        .map(|(i, _)| i)
        .unwrap_or(s.len())
}

/// Convert byte offset to char index
pub fn byte_to_char(s: &str, byte_idx: usize) -> usize {
    s[..byte_idx.min(s.len())].chars().count()
}

// ============================================================================
// Context Analyzer
// ============================================================================

/// Internal token for the analyzer (with char-based spans)
#[derive(Debug, Clone)]
struct AnalyzerToken {
    kind: AnalyzerTokenKind,
    start: usize,  // Char index
    end: usize,    // Char index (exclusive)
    text: String,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum AnalyzerTokenKind {
    Equals,
    Identifier,
    Number,
    String,
    CellRef,
    Operator,
    Comparison,
    LParen,
    RParen,
    Comma,
    Colon,
    Whitespace,
    Unknown,
}

/// Tokenize formula for analysis (produces char-indexed spans)
fn tokenize_for_analysis(formula: &str) -> Vec<AnalyzerToken> {
    let mut tokens = Vec::new();
    let chars: Vec<char> = formula.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        let start = i;
        let c = chars[i];

        match c {
            '=' if i == 0 => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Equals,
                    start,
                    end: i + 1,
                    text: "=".to_string(),
                });
                i += 1;
            }
            ' ' | '\t' => {
                // Collect whitespace
                while i < chars.len() && (chars[i] == ' ' || chars[i] == '\t') {
                    i += 1;
                }
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Whitespace,
                    start,
                    end: i,
                    text: chars[start..i].iter().collect(),
                });
            }
            '+' | '*' | '/' | '&' => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Operator,
                    start,
                    end: i + 1,
                    text: c.to_string(),
                });
                i += 1;
            }
            '-' => {
                // Could be operator or unary minus - classify based on previous token
                let is_unary = tokens.iter().rev()
                    .find(|t| t.kind != AnalyzerTokenKind::Whitespace)
                    .map(|t| matches!(t.kind,
                        AnalyzerTokenKind::Equals |
                        AnalyzerTokenKind::LParen |
                        AnalyzerTokenKind::Comma |
                        AnalyzerTokenKind::Operator |
                        AnalyzerTokenKind::Comparison
                    ))
                    .unwrap_or(true);

                tokens.push(AnalyzerToken {
                    kind: if is_unary { AnalyzerTokenKind::Operator } else { AnalyzerTokenKind::Operator },
                    start,
                    end: i + 1,
                    text: "-".to_string(),
                });
                i += 1;
            }
            '<' => {
                i += 1;
                if i < chars.len() && (chars[i] == '=' || chars[i] == '>') {
                    let text: String = chars[start..=i].iter().collect();
                    tokens.push(AnalyzerToken {
                        kind: AnalyzerTokenKind::Comparison,
                        start,
                        end: i + 1,
                        text,
                    });
                    i += 1;
                } else {
                    tokens.push(AnalyzerToken {
                        kind: AnalyzerTokenKind::Comparison,
                        start,
                        end: i,
                        text: "<".to_string(),
                    });
                }
            }
            '>' => {
                i += 1;
                if i < chars.len() && chars[i] == '=' {
                    tokens.push(AnalyzerToken {
                        kind: AnalyzerTokenKind::Comparison,
                        start,
                        end: i + 1,
                        text: ">=".to_string(),
                    });
                    i += 1;
                } else {
                    tokens.push(AnalyzerToken {
                        kind: AnalyzerTokenKind::Comparison,
                        start,
                        end: i,
                        text: ">".to_string(),
                    });
                }
            }
            '=' if i > 0 => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Comparison,
                    start,
                    end: i + 1,
                    text: "=".to_string(),
                });
                i += 1;
            }
            '(' => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::LParen,
                    start,
                    end: i + 1,
                    text: "(".to_string(),
                });
                i += 1;
            }
            ')' => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::RParen,
                    start,
                    end: i + 1,
                    text: ")".to_string(),
                });
                i += 1;
            }
            ',' => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Comma,
                    start,
                    end: i + 1,
                    text: ",".to_string(),
                });
                i += 1;
            }
            ':' => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Colon,
                    start,
                    end: i + 1,
                    text: ":".to_string(),
                });
                i += 1;
            }
            '"' => {
                // String literal
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    i += 1;
                }
                if i < chars.len() {
                    i += 1; // Include closing quote
                }
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::String,
                    start,
                    end: i,
                    text: chars[start..i].iter().collect(),
                });
            }
            '0'..='9' | '.' => {
                // Number
                while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                    i += 1;
                }
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Number,
                    start,
                    end: i,
                    text: chars[start..i].iter().collect(),
                });
            }
            'A'..='Z' | 'a'..='z' | '_' | '$' => {
                // Identifier or cell reference
                while i < chars.len() && (chars[i].is_ascii_alphanumeric() || chars[i] == '_' || chars[i] == '$') {
                    i += 1;
                }
                let text: String = chars[start..i].iter().collect();

                // Determine if it's a cell reference
                let kind = if is_cell_ref(&text) {
                    AnalyzerTokenKind::CellRef
                } else {
                    AnalyzerTokenKind::Identifier
                };

                tokens.push(AnalyzerToken {
                    kind,
                    start,
                    end: i,
                    text,
                });
            }
            _ => {
                tokens.push(AnalyzerToken {
                    kind: AnalyzerTokenKind::Unknown,
                    start,
                    end: i + 1,
                    text: c.to_string(),
                });
                i += 1;
            }
        }
    }

    tokens
}

/// Check if a string is a cell reference (e.g., A1, $B$2, AA10)
fn is_cell_ref(s: &str) -> bool {
    let s = s.to_ascii_uppercase();
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;

    // Skip leading $
    if i < chars.len() && chars[i] == '$' {
        i += 1;
    }

    // Need at least one letter
    if i >= chars.len() || !chars[i].is_ascii_uppercase() {
        return false;
    }

    // Collect letters
    while i < chars.len() && chars[i].is_ascii_uppercase() {
        i += 1;
    }

    // Skip $ before row
    if i < chars.len() && chars[i] == '$' {
        i += 1;
    }

    // Need at least one digit
    if i >= chars.len() || !chars[i].is_ascii_digit() {
        return false;
    }

    // Rest must be digits
    while i < chars.len() {
        if !chars[i].is_ascii_digit() {
            return false;
        }
        i += 1;
    }

    true
}

/// Analyze formula at cursor position
///
/// This is the single source of truth for all formula UI features.
pub fn analyze(formula: &str, cursor: usize) -> FormulaContext {
    let cursor = cursor.min(formula.chars().count());
    let tokens = tokenize_for_analysis(formula);

    // Find token at cursor
    let token_at_cursor = tokens.iter()
        .find(|t| t.start <= cursor && cursor <= t.end)
        .or_else(|| tokens.last())
        .cloned();

    // Determine mode based on context
    let mode = determine_mode(&tokens, cursor, &token_at_cursor);

    // Find current function and arg index
    let (current_function, current_arg_index, paren_depth) = find_function_context(&tokens, cursor);

    // Compute replace range
    let replace_range = compute_replace_range(&tokens, cursor, &mode);

    // Extract identifier text for autocomplete filtering
    let identifier_text = token_at_cursor.as_ref()
        .filter(|t| matches!(t.kind, AnalyzerTokenKind::Identifier))
        .map(|t| t.text.clone());

    // Build primary span
    let primary_span = token_at_cursor.as_ref()
        .filter(|t| !matches!(t.kind, AnalyzerTokenKind::Whitespace | AnalyzerTokenKind::Equals))
        .map(|t| t.start..t.end);

    // Convert token to TokenSpan
    let token_span = token_at_cursor.map(|t| TokenSpan {
        token_type: match t.kind {
            AnalyzerTokenKind::Identifier => {
                if get_function(&t.text).is_some() {
                    TokenType::Function
                } else {
                    TokenType::NamedRange // Or could be partial function name
                }
            }
            AnalyzerTokenKind::CellRef => TokenType::CellRef,
            AnalyzerTokenKind::Number => TokenType::Number,
            AnalyzerTokenKind::String => TokenType::String,
            AnalyzerTokenKind::Operator => TokenType::Operator,
            AnalyzerTokenKind::Comparison => TokenType::Comparison,
            AnalyzerTokenKind::LParen | AnalyzerTokenKind::RParen => TokenType::Paren,
            AnalyzerTokenKind::Comma => TokenType::Comma,
            AnalyzerTokenKind::Colon => TokenType::Colon,
            AnalyzerTokenKind::Whitespace => TokenType::Whitespace,
            _ => TokenType::Error,
        },
        range: t.start..t.end,
        text: t.text,
    });

    FormulaContext {
        mode,
        cursor,
        current_function,
        current_arg_index,
        token_at_cursor: token_span,
        primary_span,
        replace_range,
        paren_depth,
        identifier_text,
    }
}

fn determine_mode(tokens: &[AnalyzerToken], cursor: usize, token_at_cursor: &Option<AnalyzerToken>) -> FormulaEditMode {
    // Special case: empty formula or just "="
    if tokens.is_empty() {
        return FormulaEditMode::Start;
    }

    if tokens.len() == 1 && tokens[0].kind == AnalyzerTokenKind::Equals {
        return FormulaEditMode::Start;
    }

    // Check cursor position relative to token
    if let Some(token) = token_at_cursor {
        match token.kind {
            AnalyzerTokenKind::String => return FormulaEditMode::String,
            AnalyzerTokenKind::Number => return FormulaEditMode::Number,
            AnalyzerTokenKind::CellRef => return FormulaEditMode::Reference,
            AnalyzerTokenKind::Identifier => return FormulaEditMode::Identifier,
            AnalyzerTokenKind::Operator | AnalyzerTokenKind::Comparison => return FormulaEditMode::Operator,
            AnalyzerTokenKind::RParen => return FormulaEditMode::Complete,
            AnalyzerTokenKind::LParen | AnalyzerTokenKind::Comma => {
                // Check if we're right after it (new operand position)
                if cursor > token.start {
                    return FormulaEditMode::ArgList;
                }
            }
            _ => {}
        }
    }

    // Check what's before cursor
    let prev_token = tokens.iter()
        .filter(|t| t.kind != AnalyzerTokenKind::Whitespace)
        .filter(|t| t.end <= cursor)
        .last();

    match prev_token.map(|t| t.kind) {
        Some(AnalyzerTokenKind::Equals) => FormulaEditMode::Start,
        Some(AnalyzerTokenKind::LParen) | Some(AnalyzerTokenKind::Comma) => FormulaEditMode::ArgList,
        Some(AnalyzerTokenKind::Operator) | Some(AnalyzerTokenKind::Comparison) => FormulaEditMode::Operator,
        Some(AnalyzerTokenKind::RParen) => FormulaEditMode::Complete,
        Some(AnalyzerTokenKind::Number) | Some(AnalyzerTokenKind::CellRef) => FormulaEditMode::Complete,
        Some(AnalyzerTokenKind::Identifier) => {
            // After identifier: could be complete or waiting for (
            FormulaEditMode::Complete
        }
        _ => FormulaEditMode::Start,
    }
}

fn find_function_context(tokens: &[AnalyzerToken], cursor: usize) -> (Option<&'static FunctionInfo>, Option<usize>, usize) {
    let mut paren_depth = 0;
    let mut function_stack: Vec<(&'static FunctionInfo, usize, usize)> = Vec::new(); // (func, arg_idx, start_pos)

    for (idx, token) in tokens.iter().enumerate() {
        if token.start > cursor {
            break;
        }

        match token.kind {
            AnalyzerTokenKind::LParen => {
                // Check if previous non-whitespace token is an identifier
                let prev = tokens[..idx].iter()
                    .filter(|t| t.kind != AnalyzerTokenKind::Whitespace)
                    .last();

                if let Some(prev_tok) = prev {
                    if prev_tok.kind == AnalyzerTokenKind::Identifier {
                        if let Some(func) = get_function(&prev_tok.text) {
                            function_stack.push((func, 0, token.start));
                        }
                    }
                }
                paren_depth += 1;
            }
            AnalyzerTokenKind::RParen => {
                if !function_stack.is_empty() && token.end <= cursor {
                    function_stack.pop();
                }
                if paren_depth > 0 {
                    paren_depth -= 1;
                }
            }
            AnalyzerTokenKind::Comma => {
                if let Some((_, arg_idx, _)) = function_stack.last_mut() {
                    if token.end <= cursor {
                        *arg_idx += 1;
                    }
                }
            }
            _ => {}
        }
    }

    // Return innermost function at cursor
    if let Some((func, arg_idx, _)) = function_stack.last() {
        (Some(*func), Some(*arg_idx), paren_depth)
    } else {
        (None, None, paren_depth)
    }
}

fn compute_replace_range(tokens: &[AnalyzerToken], cursor: usize, mode: &FormulaEditMode) -> Range<usize> {
    match mode {
        FormulaEditMode::Start | FormulaEditMode::Operator | FormulaEditMode::ArgList => {
            cursor..cursor
        }
        FormulaEditMode::Identifier => {
            // Find the identifier token containing cursor
            if let Some(token) = tokens.iter().find(|t| t.kind == AnalyzerTokenKind::Identifier && t.start <= cursor && cursor <= t.end) {
                token.start..token.end
            } else {
                cursor..cursor
            }
        }
        _ => cursor..cursor,
    }
}

// ============================================================================
// Error Detection
// ============================================================================

/// Error severity for the error banner
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagnosticKind {
    /// Always show (after debounce) - user made a definite mistake
    Hard,
    /// Don't show - user is likely still typing
    Transient,
}

/// A diagnostic message
#[derive(Debug, Clone)]
pub struct Diagnostic {
    pub kind: DiagnosticKind,
    pub message: String,
    pub span: Option<Range<usize>>,
}

/// Check formula for errors
pub fn check_errors(formula: &str, cursor: usize) -> Option<Diagnostic> {
    let cursor = cursor.min(formula.chars().count());
    let formula_len = formula.chars().count();
    let tokens = tokenize_for_analysis(formula);
    let at_end = cursor >= formula_len;

    // Check for unknown tokens
    for token in &tokens {
        if token.kind == AnalyzerTokenKind::Unknown {
            return Some(Diagnostic {
                kind: DiagnosticKind::Hard,
                message: format!("Invalid character: '{}'", token.text),
                span: Some(token.start..token.end),
            });
        }
    }

    // Check for unknown functions
    for token in &tokens {
        if token.kind == AnalyzerTokenKind::Identifier {
            // Check if followed by (
            let next = tokens.iter()
                .filter(|t| t.kind != AnalyzerTokenKind::Whitespace)
                .find(|t| t.start >= token.end);

            if let Some(next_tok) = next {
                if next_tok.kind == AnalyzerTokenKind::LParen {
                    if get_function(&token.text).is_none() {
                        return Some(Diagnostic {
                            kind: DiagnosticKind::Hard,
                            message: format!("Unknown function: {}", token.text.to_uppercase()),
                            span: Some(token.start..token.end),
                        });
                    }
                }
            }
        }
    }

    // Check for unmatched parentheses
    let mut depth = 0;
    for token in &tokens {
        match token.kind {
            AnalyzerTokenKind::LParen => depth += 1,
            AnalyzerTokenKind::RParen => {
                if depth == 0 {
                    return Some(Diagnostic {
                        kind: DiagnosticKind::Hard,
                        message: "Unexpected closing parenthesis".to_string(),
                        span: Some(token.start..token.end),
                    });
                }
                depth -= 1;
            }
            _ => {}
        }
    }

    if depth > 0 {
        let kind = if at_end { DiagnosticKind::Transient } else { DiagnosticKind::Hard };
        return Some(Diagnostic {
            kind,
            message: "Missing closing parenthesis".to_string(),
            span: None,
        });
    }

    // Check for trailing operators
    let last_meaningful = tokens.iter()
        .filter(|t| !matches!(t.kind, AnalyzerTokenKind::Whitespace))
        .last();

    if let Some(last) = last_meaningful {
        if matches!(last.kind, AnalyzerTokenKind::Operator | AnalyzerTokenKind::Comparison) {
            let kind = if at_end { DiagnosticKind::Transient } else { DiagnosticKind::Hard };
            return Some(Diagnostic {
                kind,
                message: "Expected operand after operator".to_string(),
                span: Some(last.start..last.end),
            });
        }
    }

    None
}

// ============================================================================
// Tokenize for Highlighting
// ============================================================================

/// Tokenize formula and return spans for syntax highlighting
pub fn tokenize_for_highlight(formula: &str) -> Vec<(Range<usize>, TokenType)> {
    let tokens = tokenize_for_analysis(formula);

    tokens.into_iter()
        .filter(|t| t.kind != AnalyzerTokenKind::Whitespace)
        .map(|t| {
            let token_type = match t.kind {
                AnalyzerTokenKind::Equals => TokenType::Operator,
                AnalyzerTokenKind::Identifier => {
                    let upper = t.text.to_ascii_uppercase();
                    if upper == "TRUE" || upper == "FALSE" {
                        TokenType::Boolean
                    } else if get_function(&t.text).is_some() {
                        TokenType::Function
                    } else {
                        TokenType::NamedRange
                    }
                }
                AnalyzerTokenKind::Number => TokenType::Number,
                AnalyzerTokenKind::String => TokenType::String,
                AnalyzerTokenKind::CellRef => TokenType::CellRef,
                AnalyzerTokenKind::Operator => TokenType::Operator,
                AnalyzerTokenKind::Comparison => TokenType::Comparison,
                AnalyzerTokenKind::LParen | AnalyzerTokenKind::RParen => TokenType::Paren,
                AnalyzerTokenKind::Comma => TokenType::Comma,
                AnalyzerTokenKind::Colon => TokenType::Colon,
                _ => TokenType::Error,
            };
            (t.start..t.end, token_type)
        })
        .collect()
}

// ============================================================================
// Tests
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_analyze_start() {
        let ctx = analyze("=", 1);
        assert_eq!(ctx.mode, FormulaEditMode::Start);
    }

    #[test]
    fn test_analyze_identifier() {
        let ctx = analyze("=SU", 3);
        assert_eq!(ctx.mode, FormulaEditMode::Identifier);
        assert_eq!(ctx.identifier_text, Some("SU".to_string()));
    }

    #[test]
    fn test_analyze_identifier_lowercase() {
        let ctx = analyze("=su", 3);
        assert_eq!(ctx.mode, FormulaEditMode::Identifier);
        assert_eq!(ctx.identifier_text, Some("su".to_string()));
        // Should match SUM when looking up functions
        let funcs = get_functions_by_prefix("su");
        assert!(!funcs.is_empty());
        assert!(funcs.iter().any(|f| f.name == "SUM"));
    }

    #[test]
    fn test_analyze_string() {
        let ctx = analyze("=\"hello\"", 5);
        assert_eq!(ctx.mode, FormulaEditMode::String);
    }

    #[test]
    fn test_analyze_function_context() {
        let ctx = analyze("=SUM(A1,", 8);
        assert!(ctx.current_function.is_some());
        assert_eq!(ctx.current_function.unwrap().name, "SUM");
        assert_eq!(ctx.current_arg_index, Some(1));
    }

    #[test]
    fn test_analyze_nested_function() {
        let ctx = analyze("=IF(SUM(", 8);
        assert!(ctx.current_function.is_some());
        assert_eq!(ctx.current_function.unwrap().name, "SUM");
        assert_eq!(ctx.current_arg_index, Some(0));
    }

    #[test]
    fn test_is_cell_ref() {
        assert!(is_cell_ref("A1"));
        assert!(is_cell_ref("$A$1"));
        assert!(is_cell_ref("AA10"));
        assert!(is_cell_ref("$AA$100"));
        assert!(!is_cell_ref("SUM"));
        assert!(!is_cell_ref("A"));
        assert!(!is_cell_ref("1A"));
    }

    #[test]
    fn test_check_errors_unmatched_paren() {
        let diag = check_errors("=SUM(A1", 7);
        assert!(diag.is_some());
        let d = diag.unwrap();
        assert_eq!(d.kind, DiagnosticKind::Transient); // At end
        assert!(d.message.contains("parenthesis"));
    }

    #[test]
    fn test_check_errors_unknown_function() {
        let diag = check_errors("=SUMM(A1)", 5);
        assert!(diag.is_some());
        let d = diag.unwrap();
        assert_eq!(d.kind, DiagnosticKind::Hard);
        assert!(d.message.contains("SUMM"));
    }

    #[test]
    fn test_get_functions_by_prefix() {
        let funcs = get_functions_by_prefix("SU");
        assert!(funcs.iter().any(|f| f.name == "SUM"));
        assert!(funcs.iter().any(|f| f.name == "SUMIF"));
        assert!(funcs.iter().any(|f| f.name == "SUBSTITUTE"));
    }

    // =========================================================================
    // Regression tests for formula reference highlighting (Phase 2.1)
    // =========================================================================

    #[test]
    fn test_unicode_sheet_name_no_panic() {
        // Unicode sheet name: should tokenize without panic
        // Char indices must be correct even with multi-byte chars
        let formula = "='Ünicode'!A1 + B2";
        let tokens = tokenize_for_highlight(formula);

        // Should have tokens (function of formula structure)
        assert!(!tokens.is_empty(), "Should produce tokens");

        // Verify char→byte conversion works
        for (range, _) in &tokens {
            let byte_start = char_to_byte(formula, range.start);
            let byte_end = char_to_byte(formula, range.end);
            // This should not panic - proves char indices are valid
            let _slice = &formula[byte_start..byte_end];
        }

        // Find cell refs - should have A1 and B2
        let cell_refs: Vec<_> = tokens.iter()
            .filter(|(_, t)| *t == TokenType::CellRef)
            .collect();
        assert_eq!(cell_refs.len(), 2, "Should find A1 and B2");
    }

    #[test]
    fn test_duplicate_ref_same_token_content() {
        // Duplicate ref: =A1+A1 - both A1 refs should have same text content
        let formula = "=A1+A1";
        let tokens = tokenize_for_highlight(formula);

        let cell_refs: Vec<_> = tokens.iter()
            .filter(|(_, t)| *t == TokenType::CellRef)
            .collect();

        assert_eq!(cell_refs.len(), 2, "Should find two A1 refs");

        // Both should extract to "A1"
        for (range, _) in &cell_refs {
            let byte_start = char_to_byte(formula, range.start);
            let byte_end = char_to_byte(formula, range.end);
            let text = &formula[byte_start..byte_end];
            assert_eq!(text, "A1", "Both refs should be A1");
        }

        // Verify they have different positions
        assert_ne!(cell_refs[0].0.start, cell_refs[1].0.start, "Refs at different positions");
    }

    #[test]
    fn test_two_ranges_distinct_tokens() {
        // Two ranges: =SUM(A1:B5,C1:D5) - both ranges tokenized correctly
        let formula = "=SUM(A1:B5,C1:D5)";
        let tokens = tokenize_for_highlight(formula);

        // Find all cell refs (A1, B5, C1, D5)
        let cell_refs: Vec<_> = tokens.iter()
            .filter(|(_, t)| *t == TokenType::CellRef)
            .collect();
        assert_eq!(cell_refs.len(), 4, "Should find 4 cell refs: A1, B5, C1, D5");

        // Find colons
        let colons: Vec<_> = tokens.iter()
            .filter(|(_, t)| *t == TokenType::Colon)
            .collect();
        assert_eq!(colons.len(), 2, "Should find 2 colons for ranges");

        // Verify extraction works for all tokens
        for (range, _) in &tokens {
            let byte_start = char_to_byte(formula, range.start);
            let byte_end = char_to_byte(formula, range.end);
            // Should not panic
            let _text = &formula[byte_start..byte_end];
        }

        // Verify the cell ref texts
        let ref_texts: Vec<_> = cell_refs.iter()
            .map(|(range, _)| {
                let byte_start = char_to_byte(formula, range.start);
                let byte_end = char_to_byte(formula, range.end);
                &formula[byte_start..byte_end]
            })
            .collect();
        assert!(ref_texts.contains(&"A1"));
        assert!(ref_texts.contains(&"B5"));
        assert!(ref_texts.contains(&"C1"));
        assert!(ref_texts.contains(&"D5"));
    }

    #[test]
    fn test_char_to_byte_with_unicode() {
        // Direct test of char→byte conversion
        let s = "=Ü+A1";  // Ü is 2 bytes in UTF-8

        // Char indices: = is 0, Ü is 1, + is 2, A is 3, 1 is 4
        // Byte indices: = is 0, Ü is 1-2 (2 bytes), + is 3, A is 4, 1 is 5

        assert_eq!(char_to_byte(s, 0), 0);  // '='
        assert_eq!(char_to_byte(s, 1), 1);  // 'Ü' starts at byte 1
        assert_eq!(char_to_byte(s, 2), 3);  // '+' starts at byte 3 (after 2-byte Ü)
        assert_eq!(char_to_byte(s, 3), 4);  // 'A'
        assert_eq!(char_to_byte(s, 4), 5);  // '1'
        assert_eq!(char_to_byte(s, 5), 6);  // end of string
    }
}
