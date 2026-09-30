// Regular-expression functions: REGEXTEST (and REGEXMATCH, its Google Sheets
// name), REGEXEXTRACT, REGEXREPLACE.
//
// Signatures follow Excel's, which accept everything a Sheets formula passes.
// Excel's engine is PCRE2; this one is Rust's `regex`, which runs in linear time
// and so has no lookaround or backreferences. A pattern that uses them fails to
// compile and returns #VALUE!, rather than being matched differently.

use std::cell::RefCell;
use std::collections::HashMap;

use regex::{Regex, RegexBuilder};

use super::eval::{evaluate, Array2D, CellLookup, EvalResult, Value};
use super::parser::BoundExpr;

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str, args: &[BoundExpr], lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "REGEXTEST" | "REGEXMATCH" => {
            // REGEXTEST(text, pattern, [case_sensitivity])
            if args.len() < 2 || args.len() > 3 {
                return Some(EvalResult::Error(format!("{name} requires 2 or 3 arguments")));
            }
            let (text, re) = match text_and_pattern(args, 2, lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            EvalResult::Boolean(re.is_match(&text))
        }
        "REGEXEXTRACT" => {
            // REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])
            //   0 (default) the first match; 1 every match; 2 the first match's
            //   capture groups. 1 and 2 spill across a row. No match is #N/A.
            if args.len() < 2 || args.len() > 4 {
                return Some(EvalResult::Error("REGEXEXTRACT requires 2 to 4 arguments".to_string()));
            }
            let mode = match optional_int(args, 2, 0, lookup) {
                Ok(m @ 0..=2) => m,
                Ok(_) => return Some(EvalResult::Error("#VALUE! return_mode must be 0, 1 or 2".to_string())),
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let (text, re) = match text_and_pattern(args, 3, lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let found: Vec<String> = match mode {
                0 => re.find(&text).map(|m| m.as_str().to_string()).into_iter().collect(),
                1 => re.find_iter(&text).map(|m| m.as_str().to_string()).collect(),
                _ => match re.captures(&text) {
                    // A pattern without groups returns the whole match, as Excel does.
                    Some(c) if c.len() == 1 => vec![c[0].to_string()],
                    Some(c) => c.iter().skip(1).map(|g| g.map_or(String::new(), |g| g.as_str().to_string())).collect(),
                    None => Vec::new(),
                },
            };
            match found.len() {
                0 => EvalResult::Error("#N/A".to_string()),
                1 if mode == 0 => EvalResult::Text(found.into_iter().next().unwrap()),
                n => {
                    let mut row = Array2D::new(1, n);
                    for (i, s) in found.into_iter().enumerate() {
                        row.set(0, i, Value::Text(s));
                    }
                    EvalResult::Array(row)
                }
            }
        }
        "REGEXREPLACE" => {
            // REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])
            //   occurrence 0 (default) replaces every match; n replaces the nth;
            //   -n the nth from the end. $1 in the replacement is group 1.
            if args.len() < 3 || args.len() > 5 {
                return Some(EvalResult::Error("REGEXREPLACE requires 3 to 5 arguments".to_string()));
            }
            let replacement = match evaluate(&args[2], lookup) {
                EvalResult::Error(e) => return Some(EvalResult::Error(e)),
                other => excel_replacement(&other.to_text()),
            };
            let occurrence = match optional_int(args, 3, 0, lookup) {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let (text, re) = match text_and_pattern(args, 4, lookup) {
                Ok(v) => v,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            EvalResult::Text(replace_occurrence(&re, &text, &replacement, occurrence))
        }
        _ => return None,
    };
    Some(result)
}

/// The text (argument 0) and the compiled pattern (argument 1), honouring the
/// case_sensitivity argument at `case_arg` (0, the default, is case sensitive).
fn text_and_pattern<L: CellLookup>(
    args: &[BoundExpr],
    case_arg: usize,
    lookup: &L,
) -> Result<(String, Regex), String> {
    let text = match evaluate(&args[0], lookup) {
        EvalResult::Error(e) => return Err(e),
        other => other.to_text(),
    };
    let pattern = match evaluate(&args[1], lookup) {
        EvalResult::Error(e) => return Err(e),
        other => other.to_text(),
    };
    let insensitive = match optional_int(args, case_arg, 0, lookup)? {
        0 => false,
        1 => true,
        _ => return Err("#VALUE! case_sensitivity must be 0 or 1".to_string()),
    };
    Ok((text, compiled(&pattern, insensitive)?))
}

/// An optional whole-number argument, or `default` when it is absent or empty.
fn optional_int<L: CellLookup>(args: &[BoundExpr], i: usize, default: i64, lookup: &L) -> Result<i64, String> {
    match args.get(i).map(|a| evaluate(a, lookup)) {
        None | Some(EvalResult::Empty) => Ok(default),
        Some(v) => v.to_number().map(|n| n.trunc() as i64),
    }
}

/// Compile a pattern, reusing recent compilations.
///
/// A column of =REGEXTEST(A2,"...") evaluates the same pattern once per row,
/// and compiling costs far more than matching a cell's worth of text.
fn compiled(pattern: &str, insensitive: bool) -> Result<Regex, String> {
    const LIMIT: usize = 64;
    thread_local! {
        static CACHE: RefCell<HashMap<(String, bool), Regex>> = RefCell::new(HashMap::new());
    }
    CACHE.with(|cache| {
        let key = (pattern.to_string(), insensitive);
        if let Some(re) = cache.borrow().get(&key) {
            return Ok(re.clone());
        }
        let re = RegexBuilder::new(pattern)
            .case_insensitive(insensitive)
            // Bound what a hostile pattern can make the engine build.
            .size_limit(1 << 20)
            .build()
            .map_err(|_| "#VALUE! Invalid regular expression".to_string())?;
        let mut cache = cache.borrow_mut();
        if cache.len() >= LIMIT {
            cache.clear();
        }
        cache.insert(key, re.clone());
        Ok(re)
    })
}

/// Rewrite an Excel replacement string into the `regex` crate's syntax.
///
/// Both use $1 for a group, but `regex` reads "$1a" as a group named "1a" —
/// which does not exist, so it silently becomes empty — where Excel means group
/// 1 followed by "a". Numbered groups are braced, and a `$` that introduces no
/// group is kept as a literal dollar sign.
fn excel_replacement(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 8);
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        if c != '$' {
            out.push(c);
            continue;
        }
        match chars.peek() {
            Some(d) if d.is_ascii_digit() => {
                let mut n = String::new();
                while let Some(&d) = chars.peek().filter(|d| d.is_ascii_digit()) {
                    n.push(d);
                    chars.next();
                }
                out.push_str(&format!("${{{n}}}"));
            }
            Some('{') => out.push('$'),
            Some('$') => {
                chars.next();
                out.push_str("$$");
            }
            _ => out.push_str("$$"),
        }
    }
    out
}

/// Replace every match (occurrence 0), the nth (n > 0), or the nth from the end
/// (n < 0). An occurrence past the number of matches leaves the text unchanged.
fn replace_occurrence(re: &Regex, text: &str, replacement: &str, occurrence: i64) -> String {
    if occurrence == 0 {
        return re.replace_all(text, replacement).into_owned();
    }
    let matches: Vec<regex::Captures> = re.captures_iter(text).collect();
    let n = matches.len() as i64;
    let index = if occurrence > 0 { occurrence - 1 } else { n + occurrence };
    let Some(caps) = (0..n).contains(&index).then(|| &matches[index as usize]) else {
        return text.to_string();
    };
    let whole = caps.get(0).unwrap();
    let mut expanded = String::new();
    caps.expand(replacement, &mut expanded);
    format!("{}{}{}", &text[..whole.start()], expanded, &text[whole.end()..])
}

#[cfg(test)]
mod tests {
    use crate::formula::eval::{evaluate, CellLookup, EvalResult, Value};
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
            other => panic!("{formula}: expected text, got {other:?}"),
        }
    }

    fn row(formula: &str) -> Vec<String> {
        match eval(formula) {
            EvalResult::Array(a) => {
                assert_eq!(a.rows(), 1, "{formula}: spills across a row");
                (0..a.cols())
                    .map(|c| match a.get(0, c) {
                        Some(Value::Text(t)) => t.clone(),
                        other => panic!("{formula}: {other:?}"),
                    })
                    .collect()
            }
            other => panic!("{formula}: expected an array, got {other:?}"),
        }
    }

    #[test]
    fn regextest_and_its_sheets_name() {
        assert_eq!(eval(r#"=REGEXTEST("Order 1234","\d{4}")"#), EvalResult::Boolean(true));
        assert_eq!(eval(r#"=REGEXTEST("Order","\d")"#), EvalResult::Boolean(false));
        assert_eq!(eval(r#"=REGEXMATCH("abc","B")"#), EvalResult::Boolean(false));
        assert_eq!(eval(r#"=REGEXMATCH("abc","B",1)"#), EvalResult::Boolean(true));
    }

    #[test]
    fn regexextract_modes() {
        assert_eq!(text(r#"=REGEXEXTRACT("Call 555-1234 or 555-9876","\d{3}-\d{4}")"#), "555-1234");
        assert_eq!(row(r#"=REGEXEXTRACT("Call 555-1234 or 555-9876","\d{3}-\d{4}",1)"#), ["555-1234", "555-9876"]);
        assert_eq!(row(r#"=REGEXEXTRACT("SoniaBrown","([A-Z][a-z]+)([A-Z][a-z]+)",2)"#), ["Sonia", "Brown"]);
        assert_eq!(eval(r#"=REGEXEXTRACT("no digits","\d+")"#), EvalResult::Error("#N/A".to_string()));
        assert_eq!(text(r#"=REGEXEXTRACT("ABC","b",0,1)"#), "B");
    }

    #[test]
    fn regexreplace_occurrences_and_groups() {
        // Microsoft's example: swap first and last names.
        assert_eq!(text(r#"=REGEXREPLACE("SoniaBrown","([A-Z][a-z]+)([A-Z][a-z]+)","$2, $1")"#), "Brown, Sonia");
        assert_eq!(text(r#"=REGEXREPLACE("a1b2c3","\d","")"#), "abc");
        assert_eq!(text(r##"=REGEXREPLACE("a1b2c3","\d","#",2)"##), "a1b#c3");
        assert_eq!(text(r##"=REGEXREPLACE("a1b2c3","\d","#",-1)"##), "a1b2c#");
        assert_eq!(text(r##"=REGEXREPLACE("a1b2c3","\d","#",9)"##), "a1b2c3");
        assert_eq!(text(r#"=REGEXREPLACE("Cat cat","cat","dog",0,1)"#), "dog dog");
    }

    #[test]
    fn replacement_dollars_mean_what_excel_means() {
        // The regex crate reads "$1x" as a group named "1x" and substitutes
        // nothing; Excel means group 1 then "x".
        assert_eq!(text(r#"=REGEXREPLACE("ab","(a)(b)","$1x$2")"#), "axb");
        // A $ that names no group is a literal dollar sign.
        assert_eq!(text(r#"=REGEXREPLACE("5","(\d)","$ $1")"#), "$ 5");
        assert_eq!(text(r#"=REGEXREPLACE("5","(\d)","$$1")"#), "$1");
    }

    #[test]
    fn unsupported_or_invalid_patterns_are_errors() {
        // Lookaround and backreferences need PCRE; this engine refuses them
        // rather than matching differently.
        assert!(matches!(eval(r#"=REGEXTEST("ab","a(?=b)")"#), EvalResult::Error(_)));
        assert!(matches!(eval(r#"=REGEXTEST("aa","(a)\1")"#), EvalResult::Error(_)));
        assert!(matches!(eval(r#"=REGEXTEST("a","(")"#), EvalResult::Error(_)));
        assert!(matches!(eval(r#"=REGEXEXTRACT("a","a",3)"#), EvalResult::Error(_)));
        assert!(matches!(eval(r#"=REGEXTEST("a","a",2)"#), EvalResult::Error(_)));
    }
}
