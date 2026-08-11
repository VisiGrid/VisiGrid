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
        // Excel counts characters, not bytes. These index by chars throughout,
        // so a name with an accent in it behaves the same as one without.
        "SEARCH" => {
            // FIND's case-insensitive twin. Excel also accepts ? and *
            // wildcards here; this does not, and a pattern using them simply
            // fails to match rather than matching something approximate.
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error("SEARCH requires 2 or 3 arguments".to_string()));
            }
            let needle = evaluate(&args[0], lookup).to_text().to_lowercase();
            let haystack = evaluate(&args[1], lookup).to_text();
            let folded: Vec<char> = haystack.to_lowercase().chars().collect();
            let needle_chars: Vec<char> = needle.chars().collect();
            let start = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                    Ok(n) => (n as usize) - 1,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                0
            };
            if start > folded.len() {
                return Some(EvalResult::Error("#VALUE!".to_string()));
            }
            let found = if needle_chars.is_empty() {
                Some(start)
            } else {
                (start..=folded.len().saturating_sub(needle_chars.len()))
                    .find(|&i| folded[i..i + needle_chars.len()] == needle_chars[..])
            };
            match found {
                Some(pos) => EvalResult::Number((pos + 1) as f64),
                None => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "REPLACE" => {
            if args.len() != 4 {
                return Some(EvalResult::Error("REPLACE requires exactly 4 arguments".to_string()));
            }
            let text: Vec<char> = evaluate(&args[0], lookup).to_text().chars().collect();
            let start = match evaluate(&args[1], lookup).to_number() {
                Ok(n) if n < 1.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => (n as usize) - 1,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let count = match evaluate(&args[2], lookup).to_number() {
                Ok(n) if n < 0.0 => return Some(EvalResult::Error("#VALUE!".to_string())),
                Ok(n) => n as usize,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let new_text = evaluate(&args[3], lookup).to_text();
            // Starting past the end appends, which is what Excel does rather
            // than erroring.
            let head: String = text.iter().take(start).collect();
            let tail: String = text.iter().skip(start.saturating_add(count)).collect();
            EvalResult::Text(format!("{head}{new_text}{tail}"))
        }
        "PROPER" => {
            if args.len() != 1 {
                return Some(EvalResult::Error("PROPER requires exactly one argument".to_string()));
            }
            let text = evaluate(&args[0], lookup).to_text();
            // A letter starts a word when what precedes it is not a letter, so
            // "o'neill" becomes "O'Neill" and "2nd place" becomes "2Nd Place",
            // both of which are what Excel produces.
            let mut out = String::with_capacity(text.len());
            let mut prev_alpha = false;
            for ch in text.chars() {
                if ch.is_alphabetic() {
                    if prev_alpha {
                        out.extend(ch.to_lowercase());
                    } else {
                        out.extend(ch.to_uppercase());
                    }
                    prev_alpha = true;
                } else {
                    out.push(ch);
                    prev_alpha = false;
                }
            }
            EvalResult::Text(out)
        }
        "EXACT" => {
            if args.len() != 2 {
                return Some(EvalResult::Error("EXACT requires exactly 2 arguments".to_string()));
            }
            let a = evaluate(&args[0], lookup).to_text();
            let b = evaluate(&args[1], lookup).to_text();
            EvalResult::Boolean(a == b)
        }
        "TEXTBEFORE" | "TEXTAFTER" => {
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error(format!("{name} requires 2 or 3 arguments")));
            }
            let text = evaluate(&args[0], lookup).to_text();
            let delimiter = evaluate(&args[1], lookup).to_text();
            let instance = if args.len() == 3 {
                match evaluate(&args[2], lookup).to_number() {
                    Ok(n) => n as i64,
                    Err(e) => return Some(EvalResult::Error(e)),
                }
            } else {
                1
            };
            if delimiter.is_empty() || instance == 0 {
                return Some(EvalResult::Error("#VALUE!".to_string()));
            }
            let hits: Vec<usize> = text.match_indices(&delimiter).map(|(i, _)| i).collect();
            // A negative instance counts from the end, as Excel does.
            let chosen = if instance > 0 {
                hits.get((instance - 1) as usize).copied()
            } else {
                let from_end = (-instance) as usize;
                hits.len().checked_sub(from_end).and_then(|i| hits.get(i).copied())
            };
            match chosen {
                None => EvalResult::Error("#N/A".to_string()),
                Some(at) if name == "TEXTBEFORE" => EvalResult::Text(text[..at].to_string()),
                Some(at) => EvalResult::Text(text[at + delimiter.len()..].to_string()),
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

#[cfg(test)]
mod newly_added_tests {
    use crate::formula::eval::{evaluate, CellLookup, EvalResult};
    use crate::formula::parser::{bind_expr_same_sheet, parse};

    struct Empty;
    impl CellLookup for Empty {
        fn get_value(&self, _r: usize, _c: usize) -> f64 { 0.0 }
        fn get_text(&self, _r: usize, _c: usize) -> String { String::new() }
    }

    fn eval(formula: &str) -> EvalResult {
        evaluate(&bind_expr_same_sheet(&parse(formula).unwrap()), &Empty)
    }

    fn text(formula: &str) -> String {
        match eval(formula) {
            EvalResult::Text(t) => t,
            other => panic!("{formula} gave {other:?}, expected text"),
        }
    }

    fn number(formula: &str) -> f64 {
        match eval(formula) {
            EvalResult::Number(n) => n,
            other => panic!("{formula} gave {other:?}, expected a number"),
        }
    }

    fn is_error(formula: &str) -> bool {
        matches!(eval(formula), EvalResult::Error(_))
    }

    /// SEARCH is FIND without the case sensitivity.
    #[test]
    fn search_ignores_case_and_counts_from_one() {
        assert_eq!(number(r#"=SEARCH("b","ABC")"#), 2.0);
        assert_eq!(number(r#"=SEARCH("B","abc")"#), 2.0);
        assert_eq!(number(r#"=SEARCH("c","abcabc",4)"#), 6.0);
        // Excel reports #VALUE! when there is no match, not zero.
        assert!(is_error(r#"=SEARCH("z","abc")"#));
        // Counted in characters, so an accent earlier in the string does not
        // shift the answer the way byte offsets would.
        assert_eq!(number(r#"=SEARCH("é","caFÉ")"#), 4.0);
    }

    #[test]
    fn replace_works_on_character_positions() {
        assert_eq!(text(r#"=REPLACE("abcdef",2,3,"XY")"#), "aXYef");
        // Zero characters is an insert.
        assert_eq!(text(r#"=REPLACE("abcdef",2,0,"XY")"#), "aXYbcdef");
        // Starting past the end appends rather than erroring.
        assert_eq!(text(r#"=REPLACE("abc",10,2,"Z")"#), "abcZ");
    }

    #[test]
    fn proper_capitalises_after_every_non_letter() {
        assert_eq!(text(r#"=PROPER("hello world")"#), "Hello World");
        assert_eq!(text(r#"=PROPER("o'neill")"#), "O'Neill");
        // Excel really does produce "2Nd" here; the digit ends the word.
        assert_eq!(text(r#"=PROPER("2nd PLACE")"#), "2Nd Place");
        assert_eq!(text(r#"=PROPER("")"#), "");
    }

    #[test]
    fn exact_is_the_case_sensitive_comparison() {
        assert_eq!(eval(r#"=EXACT("a","A")"#), EvalResult::Boolean(false));
        assert_eq!(eval(r#"=EXACT("abc","abc")"#), EvalResult::Boolean(true));
    }

    #[test]
    fn textbefore_and_textafter_take_an_instance() {
        assert_eq!(text(r#"=TEXTBEFORE("a-b-c","-")"#), "a");
        assert_eq!(text(r#"=TEXTAFTER("a-b-c","-")"#), "b-c");
        assert_eq!(text(r#"=TEXTBEFORE("a-b-c","-",2)"#), "a-b");
        // A negative instance counts back from the end.
        assert_eq!(text(r#"=TEXTAFTER("a-b-c","-",-1)"#), "c");
        // Missing delimiter is #N/A, not an empty string — an empty string
        // would be indistinguishable from a delimiter at position one.
        assert!(is_error(r#"=TEXTBEFORE("abc","-")"#));
    }
}
