// Trigonometric and logarithmic functions: SIN, COS, TAN, ASIN, ACOS, ATAN,
// ATAN2, PI, DEGREES, RADIANS, LOG, LOG10, LN, EXP

use super::eval::{evaluate, CellLookup, EvalResult};
use super::parser::BoundExpr;

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "LOG" => {
            // LOG(number, [base]) - base defaults to 10
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("LOG requires 1 or 2 arguments".to_string()));
            }
            let number = match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n <= 0.0 => return Some(EvalResult::Error("#NUM!".to_string())),
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let base = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(b) if b <= 0.0 || b == 1.0 => return Some(EvalResult::Error("#NUM!".to_string())),
                    Ok(b) => b,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                10.0
            };
            EvalResult::Number(number.log(base))
        }
        "LOG10" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("LOG10 requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n <= 0.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.log10()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "LN" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("LN requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n <= 0.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.ln()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "EXP" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("EXP requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.exp()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "PI" => {
            if !args.is_empty() {
                return Some(EvalResult::Error("PI takes no arguments".to_string()));
            }
            EvalResult::Number(std::f64::consts::PI)
        }
        "DEGREES" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("DEGREES requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.to_degrees()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "RADIANS" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("RADIANS requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.to_radians()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "SIN" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("SIN requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.sin()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "COS" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("COS requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.cos()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "TAN" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("TAN requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.tan()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ASIN" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ASIN requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < -1.0 || n > 1.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.asin()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ACOS" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ACOS requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) if n < -1.0 || n > 1.0 => EvalResult::Error("#NUM!".to_string()),
                Ok(n) => EvalResult::Number(n.acos()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ATAN" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("ATAN requires exactly one argument".to_string()));
            }
            match evaluate(&args[0], lookup).to_number() {
                Ok(n) => EvalResult::Number(n.atan()),
                Err(e) => EvalResult::Error(e),
            }
        }
        "ATAN2" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("ATAN2 requires exactly 2 arguments".to_string()));
            }
            let x = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let y = match evaluate(&args[1], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            if x == 0.0 && y == 0.0 {
                return Some(EvalResult::Error("#DIV/0!".to_string()));
            }
            EvalResult::Number(y.atan2(x))
        }
        _ => return None,
    };
    Some(result)
}
