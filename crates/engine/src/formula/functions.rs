// Built-in spreadsheet functions

use super::eval::EvalResult;

/// All supported function names, sorted alphabetically.
/// This is the single source of truth for the function list.
const FUNCTION_NAMES: &[&str] = &[
    "ABS", "ACOS", "AND", "ASIN", "ATAN", "ATAN2",
    "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "AVG",
    "CEILING", "CHOOSE", "COLUMN", "COLUMNS", "CONCAT", "CONCATENATE",
    "COS", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS",
    "CUMIPMT", "CUMPRINC",
    "DATE", "DATEDIF", "DAY", "DEGREES",
    "EDATE", "EOMONTH", "EXP",
    "FILTER", "FIND", "FLOOR", "FV",
    "HLOOKUP", "HOUR",
    "IF", "IFERROR", "IFNA", "IFS", "INDEX", "INT", "IPMT", "IRR",
    "ISBLANK", "ISERROR", "ISNA", "ISNUMBER", "ISTEXT",
    "LEFT", "LEN", "LN", "LOG", "LOG10", "LOWER",
    "MATCH", "MAX", "MEDIAN", "MID", "MIN", "MINUTE", "MOD", "MONTH",
    "NORM.S.DIST", "NORMSDIST", "NOT", "NOW", "NPV",
    "OR",
    "PI", "PMT", "POWER", "PPMT", "PRODUCT", "PV",
    "RADIANS", "RAND", "RANDBETWEEN", "REPT", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS",
    "SECOND", "SEQUENCE", "SIN", "SORT", "SPARKLINE", "SQRT", "STDEV", "SUBSTITUTE", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SWITCH",
    "TAN", "TEXT", "TEXTJOIN", "TODAY", "TRANSPOSE", "TRIM", "TRUNC",
    "UNIQUE", "UPPER",
    "VALUE", "VAR", "VLOOKUP",
    "WEEKDAY",
    "XLOOKUP",
    "YEAR",
];

/// Returns all supported function names, sorted alphabetically.
pub fn list_functions() -> &'static [&'static str] {
    FUNCTION_NAMES
}

/// Check if a function name is a known built-in function.
pub fn is_known_function(name: &str) -> bool {
    FUNCTION_NAMES.binary_search(&name).is_ok()
}

/// Check if a name is valid for a user-defined custom function.
/// Must be non-empty, start with uppercase, and contain only uppercase + digits + underscores.
pub fn is_valid_custom_function_name(name: &str) -> bool {
    let bytes = name.as_bytes();
    !bytes.is_empty()
        && bytes[0].is_ascii_uppercase()
        && bytes.iter().all(|b| b.is_ascii_uppercase() || b.is_ascii_digit() || *b == b'_')
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
