// Text functions: CONCATENATE, TEXTJOIN, LEFT, RIGHT, MID, LEN, UPPER, LOWER,
// TRIM, TEXT, VALUE, FIND, SUBSTITUTE, REPT

use super::eval::{evaluate, CellLookup, EvalResult};
use super::parser::{BoundExpr, Expr};

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "CONCATENATE" | "CONCAT" => {
            let mut result = String::new();
            for arg in args {
                result.push_str(&evaluate(arg, lookup).to_text());
            }
            EvalResult::Text(result)
        }
        "TEXTJOIN" => {
            // TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
            if args.len() < 3 {
                return Some(EvalResult::Error("TEXTJOIN requires at least 3 arguments".to_string()));
            }
            let delimiter = evaluate(&args[0], lookup).to_text();
            let ignore_empty = match evaluate(&args[1], lookup).to_bool() {
                Ok(b) => b,
                Err(_) => true, // default to TRUE
            };

            let mut parts: Vec<String> = Vec::new();

            for arg in &args[2..] {
                match arg {
                    Expr::Range { start_col, start_row, end_col, end_row, .. } => {
                        // Collect all values from range
                        let (min_row, min_col, max_row, max_col) = (
                            (*start_row).min(*end_row), (*start_col).min(*end_col),
                            (*start_row).max(*end_row), (*start_col).max(*end_col)
                        );
                        for r in min_row..=max_row {
                            for c in min_col..=max_col {
                                let text = lookup.get_text(r, c);
                                if !ignore_empty || !text.is_empty() {
                                    parts.push(text);
                                }
                            }
                        }
                    }
                    _ => {
                        let text = evaluate(arg, lookup).to_text();
                        if !ignore_empty || !text.is_empty() {
                            parts.push(text);
                        }
                    }
                }
            }

            EvalResult::Text(parts.join(&delimiter))
        }
        "LEFT" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("LEFT requires 1 or 2 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let num_chars = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };
            EvalResult::Text(text.chars().take(num_chars).collect())
        }
        "RIGHT" => {
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error("RIGHT requires 1 or 2 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let num_chars = if args.len() == 2 {
                match evaluate(&args[1], lookup).to_number() {
                    Ok(n) => n as usize,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };
            let len = text.chars().count();
            let start = len.saturating_sub(num_chars);
            EvalResult::Text(text.chars().skip(start).collect())
        }
        "MID" => {
            if args.len() != 3 {
                return Some(EvalResult::Error("MID requires exactly 3 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let start = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => (n as usize).saturating_sub(1), // 1-indexed
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let num_chars = match evaluate(&args[2], lookup).to_number() {
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            EvalResult::Text(text.chars().skip(start).take(num_chars).collect())
        }
        "LEN" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("LEN requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Number(text.chars().count() as f64)
        }
        "UPPER" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("UPPER requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Text(text.to_uppercase())
        }
        "LOWER" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("LOWER requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            EvalResult::Text(text.to_lowercase())
        }
        "TRIM" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("TRIM requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            // TRIM removes leading/trailing spaces and collapses internal spaces
            let trimmed: String = text.split_whitespace().collect::<Vec<_>>().join(" ");
            EvalResult::Text(trimmed)
        }
        "TEXT" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("TEXT requires exactly 2 arguments".to_string()));
            }
            let value = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let format = evaluate(&args[1], lookup).to_text();
            // Simple format support
            let result = if format.contains("0.") {
                let decimals = format.matches('0').count().saturating_sub(1);
                format!("{:.1$}", value, decimals)
            } else if format.contains('%') {
                format!("{}%", (value * 100.0) as i64)
            } else {
                format!("{}", value)
            };
            EvalResult::Text(result)
        }
        "VALUE" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("VALUE requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            match text.replace(',', "").trim().parse::<f64>() {
                Ok(n) => EvalResult::Number(n),
                Err(_) => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "FIND" => {
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("FIND requires 2 or 3 arguments".to_string()));
            }
            let find_text = evaluate(&args[0], lookup).to_text();
            let within_text = evaluate(&args[1], lookup).to_text();
            let start_pos = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                    Ok(n) => (n as usize).saturating_sub(1),
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0
            };
            let search_area = &within_text[start_pos.min(within_text.len())..];
            match search_area.find(&find_text) {
                Some(pos) => EvalResult::Number((pos + start_pos + 1) as f64), // 1-indexed
                None => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "SUBSTITUTE" => {
            if args.len() < 3 || args.len() > 4 {
                return Some(EvalResult::Error("SUBSTITUTE requires 3 or 4 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let old_text = evaluate(&args[1], lookup).to_text();
            let new_text = evaluate(&args[2], lookup).to_text();
            let instance = if args.len() == 4 {
                match evaluate(&args[3], lookup).to_number() {
                    Ok(n) => Some(n as usize),
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                None
            };

            let result = if let Some(n) = instance {
                // Replace only the nth instance
                let mut count = 0;
                let mut result = String::new();
                let mut remaining = text.as_str();
                while let Some(pos) = remaining.find(&old_text) {
                    count += 1;
                    if count == n {
                        result.push_str(&remaining[..pos]);
                        result.push_str(&new_text);
                        result.push_str(&remaining[pos + old_text.len()..]);
                        break;
                    } else {
                        result.push_str(&remaining[..pos + old_text.len()]);
                        remaining = &remaining[pos + old_text.len()..];
                    }
                }
                if count < n {
                    text // Not enough instances found
                } else {
                    result
                }
            } else {
                // Replace all instances
                text.replace(&old_text, &new_text)
            };
            EvalResult::Text(result)
        }
        "REPT" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("REPT requires exactly 2 arguments".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let times = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 0.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            EvalResult::Text(text.repeat(times))
        }
        _ => return None,
    };
    Some(result)
}
