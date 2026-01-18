// Built-in spreadsheet functions

use super::eval::EvalResult;

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
