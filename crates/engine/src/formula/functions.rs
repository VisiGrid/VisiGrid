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
    "NOT", "NOW", "NPV",
    "OR",
    "PI", "PMT", "POWER", "PPMT", "PRODUCT", "PV",
    "RADIANS", "RAND", "RANDBETWEEN", "REPT", "RIGHT", "ROUND", "ROW", "ROWS",
    "SECOND", "SEQUENCE", "SIN", "SORT", "SPARKLINE", "SQRT", "STDEV", "SUBSTITUTE", "SUM", "SUMIF", "SUMIFS", "SWITCH",
    "TAN", "TEXT", "TEXTJOIN", "TODAY", "TRANSPOSE", "TRIM",
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
