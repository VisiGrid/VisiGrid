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
    "DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS", "DEGREES",
    "EDATE", "EOMONTH", "EXACT", "EXP",
    "FILTER", "FIND", "FLOOR", "FV",
    "HLOOKUP", "HOUR",
    "IF", "IFERROR", "IFNA", "IFS", "INDEX", "INDIRECT", "INT", "IPMT", "IRR",
    "ISBLANK", "ISERROR", "ISNA", "ISNUMBER", "ISTEXT",
    "LEFT", "LEN", "LN", "LOG", "LOG10", "LOWER",
    "MATCH", "MAX", "MEDIAN", "MID", "MIN", "MINUTE", "MOD", "MONTH",
    "NETWORKDAYS", "NORM.S.DIST", "NORMSDIST", "NOT", "NOW", "NPER", "NPV",
    "OFFSET", "OR",
    "PI", "PMT", "POWER", "PPMT", "PRODUCT", "PROPER", "PV",
    "RADIANS", "RAND", "RANDBETWEEN", "RATE", "REGEXEXTRACT", "REGEXMATCH", "REGEXREPLACE", "REGEXTEST", "REPLACE", "REPT", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS",
    "SEARCH", "SECOND", "SEQUENCE", "SIN", "SORT", "SPARKLINE", "SQRT", "STDEV", "STDEV.P", "STDEV.S", "STDEVP", "SUBSTITUTE", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SWITCH",
    "TAN", "TEXT", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TIME", "TODAY", "TRANSPOSE", "TRIM", "TRUNC",
    "UNIQUE", "UPPER",
    "VALUE", "VAR", "VAR.P", "VAR.S", "VARP", "VLOOKUP",
    "WEEKDAY", "WORKDAY",
    "XIRR", "XLOOKUP", "XMATCH", "XNPV",
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn function_names_are_sorted() {
        // `is_known_function` uses binary_search, which is only correct on a sorted list.
        // Note: '.' (0x2E) sorts before ASCII letters, so e.g. "STDEV.P" precedes "STDEVP".
        for pair in FUNCTION_NAMES.windows(2) {
            assert!(
                pair[0] < pair[1],
                "FUNCTION_NAMES not strictly sorted: {:?} !< {:?}",
                pair[0],
                pair[1]
            );
        }
    }

    /// Every function the evaluator dispatches must be in FUNCTION_NAMES.
    ///
    /// The list is maintained by hand beside a dispatch spread over ten files, and it has
    /// drifted twice: seven names in one release, then eleven more (DAYS, EXACT, XMATCH
    /// and the rest of 0.24.0's additions) that worked in cells while autocomplete,
    /// validation and `vgrid list-functions` said they did not exist. Read the dispatch
    /// arms from source rather than trusting another hand-kept list. Top-level arms sit at
    /// exactly eight spaces; nested matches (DATEDIF's "YM" units) are deeper.
    #[test]
    fn every_dispatched_function_is_registered() {
        let sources = [
            include_str!("eval_math.rs"),
            include_str!("eval_logical.rs"),
            include_str!("eval_text.rs"),
            include_str!("eval_regex.rs"),
            include_str!("eval_conditional.rs"),
            include_str!("eval_lookup.rs"),
            include_str!("eval_financial.rs"),
            include_str!("eval_datetime.rs"),
            include_str!("eval_trig.rs"),
            include_str!("eval_statistical.rs"),
            include_str!("eval_array.rs"),
        ];
        let mut dispatched = Vec::new();
        for src in sources {
            for line in src.lines() {
                let Some(arm) = line.strip_prefix("        \"") else { continue };
                let Some((names, _)) = arm.split_once("=>") else { continue };
                for name in format!("\"{names}").split('|') {
                    dispatched.push(name.trim().trim_matches('"').to_string());
                }
            }
        }
        assert!(dispatched.len() > 100, "scan found only {} arms; has the dispatch layout changed?", dispatched.len());
        let missing: Vec<_> = dispatched.iter().filter(|n| !is_known_function(n)).collect();
        assert!(missing.is_empty(), "dispatched but not in FUNCTION_NAMES: {missing:?}");
    }

    #[test]
    fn previously_orphaned_functions_are_registered() {
        // Regression: these were callable in dispatch but missing from FUNCTION_NAMES,
        // so is_known_function() wrongly returned false (breaking autocomplete/validation).
        for name in ["DATEVALUE", "STDEV.S", "STDEV.P", "STDEVP", "VAR.S", "VAR.P", "VARP"] {
            assert!(is_known_function(name), "{name} should be a known function");
        }
    }
}
