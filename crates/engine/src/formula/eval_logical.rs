// Logical functions: IF, AND, OR, NOT, IFERROR, IFNA, ISBLANK, ISNUMBER, ISTEXT,
// ISERROR, ISNA, IFS, SWITCH, CHOOSE

use super::eval::{evaluate, CellLookup, EvalResult};
use super::parser::BoundExpr;

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "IF" => {
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("IF requires 2 or 3 arguments".to_string()));
            }
            let condition = match evaluate(&args[0], lookup).to_bool() {
                Ok(b) => b,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if condition {
                evaluate(&args[1], lookup)
            } else if args.len() == 3 {
                evaluate(&args[2], lookup)
            } else {
                EvalResult::Boolean(false)
            }
        }
        "AND" => {
            if args.is_empty() {
                return Some(EvalResult::Error("AND requires at least one argument".to_string()));
            }
            for arg in args {
                match evaluate(arg, lookup).to_bool() {
                    Ok(false) => return Some(EvalResult::Boolean(false)),
                    Err(e) => return Some(EvalResult::Error(e)),
                    _ => {}
                }
            }
            EvalResult::Boolean(true)
        }
        "OR" => {
            if args.is_empty() {
                return Some(EvalResult::Error("OR requires at least one argument".to_string()));
            }
            for arg in args {
                match evaluate(arg, lookup).to_bool() {
                    Ok(true) => return Some(EvalResult::Boolean(true)),
                    Err(e) => return Some(EvalResult::Error(e)),
                    _ => {}
                }
            }
            EvalResult::Boolean(false)
        }
        "NOT" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("NOT requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_bool() {
                Ok(b) => EvalResult::Boolean(!b),
                Err(e) => EvalResult::Error(e),
            }
        }
        "IFERROR" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("IFERROR requires exactly 2 arguments".to_string()));
            }
            let result = evaluate(&args[0], lookup);
            if result.is_error() {
                evaluate(&args[1], lookup)
            } else {
                result
            }
        }
        "ISBLANK" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ISBLANK requires exactly one argument".to_string()));
            }
            let result = evaluate(&args[0], lookup);
            let is_blank = match &result {
                EvalResult::Text(s) => s.is_empty(),
                _ => false,
            };
            EvalResult::Boolean(is_blank)
        }
        "ISNUMBER" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ISNUMBER requires exactly one argument".to_string()));
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(matches!(result, EvalResult::Number(_)))
        }
        "ISTEXT" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ISTEXT requires exactly one argument".to_string()));
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(matches!(result, EvalResult::Text(_)))
        }
        "ISERROR" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ISERROR requires exactly one argument".to_string()));
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(result.is_error())
        }
        "ISNA" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ISNA requires exactly one argument".to_string()));
            }
            let result = evaluate(&args[0], lookup);
            EvalResult::Boolean(matches!(result, EvalResult::Error(ref e) if e == "#N/A"))
        }
        "IFNA" => {
            // IFNA(value, value_if_na)
            if args.len() != 2 {
                return Some(EvalResult::Error("IFNA requires exactly 2 arguments".to_string()));
            }
            let value = evaluate(&args[0], lookup);
            match value {
                EvalResult::Error(ref e) if e == "#N/A" => evaluate(&args[1], lookup),
                _ => value,
            }
        }
        "IFS" => {
            // IFS(condition1, value1, [condition2, value2], ...)
            if args.len() < 2 || args.len() % 2 != 0 {
                return Some(EvalResult::Error("IFS requires pairs of condition, value arguments".to_string()));
            }
            for i in (0..args.len()).step_by(2) {
                match evaluate(&args[i], lookup).to_bool() {
                    Ok(true) => return Some(evaluate(&args[i + 1], lookup)),
                    Ok(false) => continue,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            }
            EvalResult::Error("#N/A".to_string())
        }
        "SWITCH" => {
            // SWITCH(expression, value1, result1, [value2, result2], ..., [default])
            if args.len() < 3 {
                return Some(EvalResult::Error("SWITCH requires at least 3 arguments".to_string()));
            }
            let expr = evaluate(&args[0], lookup);
            let pairs = (args.len() - 1) / 2;
            let has_default = (args.len() - 1) % 2 == 1;

            for i in 0..pairs {
                let value = evaluate(&args[1 + i * 2], lookup);
                // Compare expr with value
                let matches = match (&expr, &value) {
                    (EvalResult::Number(a), EvalResult::Number(b)) => (a - b).abs() < f64::EPSILON,
                    (EvalResult::Text(a), EvalResult::Text(b)) => a.to_lowercase() == b.to_lowercase(),
                    (EvalResult::Boolean(a), EvalResult::Boolean(b)) => a == b,
                    _ => expr.to_text() == value.to_text(),
                };
                if matches {
                    return Some(evaluate(&args[2 + i * 2], lookup));
                }
            }

            if has_default {
                evaluate(&args[args.len() - 1], lookup)
            } else {
                EvalResult::Error("#N/A".to_string())
            }
        }
        "CHOOSE" => {
            // CHOOSE(index, value1, [value2], ...)
            if args.len() < 2 {
                return Some(EvalResult::Error("CHOOSE requires at least 2 arguments".to_string()));
            }
            let index = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if index < 1 || index >= args.len() {
                return Some(EvalResult::Error("#VALUE!".to_string()));
            }
            evaluate(&args[index], lookup)
        }
        _ => return None,
    };
    Some(result)
}
