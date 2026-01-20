// Built-in spreadsheet functions

use super::eval::EvalResult;

/// Check if a function name is a known built-in function.
/// This is the single source of truth for supported functions.
/// Function names must be uppercase (as produced by the parser).
pub fn is_known_function(name: &str) -> bool {
    matches!(name,
        // Math functions
        "SUM" | "AVERAGE" | "AVG" | "MIN" | "MAX" | "COUNT" | "COUNTA" |
        "ABS" | "ROUND" | "INT" | "MOD" | "POWER" | "SQRT" | "CEILING" | "FLOOR" |
        "PRODUCT" | "MEDIAN" | "STDEV" | "VAR" |
        // Logic functions
        "IF" | "IFS" | "AND" | "OR" | "NOT" | "IFERROR" | "IFNA" | "SWITCH" |
        // Information functions
        "ISBLANK" | "ISNUMBER" | "ISTEXT" | "ISERROR" |
        // Text functions
        "CONCATENATE" | "CONCAT" | "TEXTJOIN" | "LEFT" | "RIGHT" | "MID" | "LEN" |
        "UPPER" | "LOWER" | "TRIM" | "TEXT" | "VALUE" | "FIND" | "SUBSTITUTE" | "REPT" |
        // Conditional aggregation
        "SUMIF" | "COUNTIF" | "COUNTBLANK" | "SUMIFS" | "COUNTIFS" |
        // Lookup functions
        "VLOOKUP" | "HLOOKUP" | "XLOOKUP" | "INDEX" | "MATCH" | "CHOOSE" |
        // Reference functions
        "ROW" | "COLUMN" | "ROWS" | "COLUMNS" |
        // Date/time functions
        "TODAY" | "NOW" | "DATE" | "YEAR" | "MONTH" | "DAY" | "WEEKDAY" |
        "DATEDIF" | "EDATE" | "EOMONTH" | "HOUR" | "MINUTE" | "SECOND" |
        // Random functions
        "RAND" | "RANDBETWEEN" |
        // Math/trig functions
        "LOG" | "LOG10" | "LN" | "EXP" | "PI" | "DEGREES" | "RADIANS" |
        "SIN" | "COS" | "TAN" | "ASIN" | "ACOS" | "ATAN" | "ATAN2" |
        // Array functions
        "FILTER" | "SORT" | "UNIQUE" | "SEQUENCE" | "TRANSPOSE"
    )
}

pub type FunctionImpl = fn(args: &[EvalResult]) -> EvalResult;

pub fn sum(args: &[EvalResult]) -> EvalResult {
    let mut total = 0.0;
    for arg in args {
        match arg {
            EvalResult::Number(n) => total += n,
            EvalResult::Error(e) => return EvalResult::Error(e.clone()),
            _ => {}
        }
    }
    EvalResult::Number(total)
}

pub fn average(args: &[EvalResult]) -> EvalResult {
    let mut total = 0.0;
    let mut count = 0;
    for arg in args {
        match arg {
            EvalResult::Number(n) => {
                total += n;
                count += 1;
            }
            EvalResult::Error(e) => return EvalResult::Error(e.clone()),
            _ => {}
        }
    }
    if count == 0 {
        EvalResult::Error("Division by zero".to_string())
    } else {
        EvalResult::Number(total / count as f64)
    }
}

pub fn min(args: &[EvalResult]) -> EvalResult {
    let mut result: Option<f64> = None;
    for arg in args {
        match arg {
            EvalResult::Number(n) => {
                result = Some(result.map_or(*n, |r| r.min(*n)));
            }
            EvalResult::Error(e) => return EvalResult::Error(e.clone()),
            _ => {}
        }
    }
    result.map(EvalResult::Number).unwrap_or(EvalResult::Number(0.0))
}

pub fn max(args: &[EvalResult]) -> EvalResult {
    let mut result: Option<f64> = None;
    for arg in args {
        match arg {
            EvalResult::Number(n) => {
                result = Some(result.map_or(*n, |r| r.max(*n)));
            }
            EvalResult::Error(e) => return EvalResult::Error(e.clone()),
            _ => {}
        }
    }
    result.map(EvalResult::Number).unwrap_or(EvalResult::Number(0.0))
}
