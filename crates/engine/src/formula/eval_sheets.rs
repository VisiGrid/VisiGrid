// Google Sheets functions with no Excel equivalent: the operator functions
// (ADD, MINUS, MULTIPLY, DIVIDE, EQ, NE, GT, GTE, LT, LTE, UMINUS, UPLUS,
// UNARY_PERCENT), the TO_* conversions, ISBETWEEN, ISEMAIL, ISURL, EPOCHTODATE
// JOIN and QUERY (eval_query.rs). They exist so formulas written in Sheets (and the ones a Sheets
// .xlsx export carries) keep working.
//
// All but JOIN take plain values and are lifted (lift.rs), so over a range they
// go element by element: =ADD(A1:A3, 1) spills, as it does in Sheets' array
// mode, which is the only mode here.

use super::eval::{evaluate, operand, scalar_binary, CellLookup, EvalResult, Value};
use super::parser::{BoundExpr, Op};

pub(crate) fn try_evaluate<L: CellLookup>(
    name: &str,
    args: &[BoundExpr],
    lookup: &L,
) -> Option<EvalResult> {
    let result = match name {
        "ADD" | "MINUS" | "MULTIPLY" | "DIVIDE" | "EQ" | "NE" | "GT" | "GTE" | "LT" | "LTE" => {
            // The operator, as a function: ADD(a, b) is a+b, EQ(a, b) is a=b,
            // with the operator's coercions and errors.
            if args.len() != 2 {
                return Some(EvalResult::Error(format!(
                    "{name} requires exactly 2 arguments"
                )));
            }
            let op = match name {
                "ADD" => Op::Add,
                "MINUS" => Op::Sub,
                "MULTIPLY" => Op::Mul,
                "DIVIDE" => Op::Div,
                "EQ" => Op::Eq,
                "NE" => Op::NotEq,
                "GT" => Op::Gt,
                "GTE" => Op::GtEq,
                "LT" => Op::Lt,
                _ => Op::LtEq,
            };
            scalar_binary(op, evaluate(&args[0], lookup), evaluate(&args[1], lookup))
        }
        "UMINUS" | "UNARY_PERCENT" => {
            // -value and value%: text that isn't a number is #VALUE!.
            if args.len() != 1 {
                return Some(EvalResult::Error(format!(
                    "{name} requires exactly 1 argument"
                )));
            }
            let value = evaluate(&args[0], lookup);
            if name == "UMINUS" {
                scalar_binary(Op::Sub, EvalResult::Number(0.0), value)
            } else {
                scalar_binary(Op::Div, value, EvalResult::Number(100.0))
            }
        }
        "UPLUS" => {
            // +value: the value itself, text included, as in Sheets. A blank is 0.
            if args.len() != 1 {
                return Some(EvalResult::Error(
                    "UPLUS requires exactly 1 argument".to_string(),
                ));
            }
            match evaluate(&args[0], lookup) {
                EvalResult::Empty => EvalResult::Number(0.0),
                other => other,
            }
        }
        "TO_TEXT" => {
            // A number as text. Sheets uses the cell's display format; here a
            // value carries no format, so numbers read as they do in General.
            if args.len() != 1 {
                return Some(EvalResult::Error(
                    "TO_TEXT requires exactly 1 argument".to_string(),
                ));
            }
            match evaluate(&args[0], lookup) {
                EvalResult::Empty => EvalResult::Text(String::new()),
                EvalResult::Error(e) => EvalResult::Error(e),
                other => EvalResult::Text(other.to_text()),
            }
        }
        "TO_DATE" | "TO_PERCENT" | "TO_DOLLARS" | "TO_PURE_NUMBER" => {
            // In Sheets these change how a number is shown, never its value,
            // and return anything that isn't a number unchanged. A result here
            // carries no number format, so the number comes back as it is.
            if args.len() != 1 {
                return Some(EvalResult::Error(format!(
                    "{name} requires exactly 1 argument"
                )));
            }
            match evaluate(&args[0], lookup) {
                EvalResult::Empty if name == "TO_PURE_NUMBER" => EvalResult::Number(0.0),
                other => other,
            }
        }
        "ISBETWEEN" => {
            // ISBETWEEN(value, lower, upper, [lower_inclusive], [upper_inclusive]):
            // both bounds are inclusive unless switched off. Values compare as
            // the comparison operators compare them.
            if args.len() < 3 || args.len() > 5 {
                return Some(EvalResult::Error(
                    "ISBETWEEN requires 3 to 5 arguments".to_string(),
                ));
            }
            let flag = |i: usize| -> Result<bool, String> {
                match args.get(i) {
                    None => Ok(true),
                    Some(arg) => evaluate(arg, lookup).to_bool(),
                }
            };
            let (lower_inclusive, upper_inclusive) = match (flag(3), flag(4)) {
                (Ok(l), Ok(u)) => (l, u),
                (Err(e), _) | (_, Err(e)) => return Some(EvalResult::Error(e)),
            };
            let value = evaluate(&args[0], lookup);
            let above = scalar_binary(
                if lower_inclusive { Op::GtEq } else { Op::Gt },
                value.clone(),
                evaluate(&args[1], lookup),
            );
            let below = scalar_binary(
                if upper_inclusive { Op::LtEq } else { Op::Lt },
                value,
                evaluate(&args[2], lookup),
            );
            match (above, below) {
                (EvalResult::Boolean(a), EvalResult::Boolean(b)) => EvalResult::Boolean(a && b),
                (EvalResult::Error(e), _) | (_, EvalResult::Error(e)) => EvalResult::Error(e),
                _ => EvalResult::Error("#VALUE!".to_string()),
            }
        }
        "ISEMAIL" | "ISURL" => {
            // TRUE for text shaped like an email address or a web address.
            // Anything else, an error included, is FALSE.
            if args.len() != 1 {
                return Some(EvalResult::Error(format!(
                    "{name} requires exactly 1 argument"
                )));
            }
            let EvalResult::Text(text) = evaluate(&args[0], lookup) else {
                return Some(EvalResult::Boolean(false));
            };
            let text = text.trim();
            EvalResult::Boolean(if name == "ISEMAIL" {
                is_email(text)
            } else {
                is_url(text)
            })
        }
        "EPOCHTODATE" => {
            // EPOCHTODATE(timestamp, [unit]): a Unix timestamp as a UTC date
            // serial. unit 1 (the default) is seconds, 2 milliseconds, 3
            // microseconds; precision stops at the millisecond, as in Sheets.
            if args.is_empty() || args.len() > 2 {
                return Some(EvalResult::Error(
                    "EPOCHTODATE requires 1 or 2 arguments".to_string(),
                ));
            }
            let timestamp = match evaluate(&args[0], lookup).to_number() {
                Ok(n) => n,
                Err(e) => return Some(EvalResult::Error(e)),
            };
            let unit = match args.get(1).map(|a| evaluate(a, lookup).to_number()) {
                None => 1.0,
                Some(Ok(n)) => n.trunc(),
                Some(Err(e)) => return Some(EvalResult::Error(e)),
            };
            let per_ms = match unit as i64 {
                1 => 1.0 / 1000.0,
                2 => 1.0,
                3 => 1000.0,
                _ => {
                    return Some(EvalResult::Error(
                        "#NUM! EPOCHTODATE unit must be 1, 2 or 3".to_string(),
                    ))
                }
            };
            if timestamp < 0.0 {
                return Some(EvalResult::Error(
                    "#NUM! EPOCHTODATE timestamp must not be negative".to_string(),
                ));
            }
            let ms = (timestamp / per_ms).trunc();
            EvalResult::Number(25569.0 + ms / 86_400_000.0)
        }
        "JOIN" => {
            // JOIN(delimiter, value_or_array1, ...): every value, blanks
            // included, joined by the delimiter. Each array must be a single
            // row or column.
            if args.len() < 2 {
                return Some(EvalResult::Error(
                    "JOIN requires a delimiter and at least one value".to_string(),
                ));
            }
            let delimiter = match evaluate(&args[0], lookup) {
                EvalResult::Error(e) => return Some(EvalResult::Error(e)),
                d => d.to_text(),
            };
            let mut parts: Vec<String> = Vec::new();
            for arg in &args[1..] {
                match operand(arg, lookup) {
                    EvalResult::Array(a) => {
                        if a.rows() > 1 && a.cols() > 1 {
                            return Some(EvalResult::Error(
                                "#VALUE! JOIN takes a single row or column".to_string(),
                            ));
                        }
                        for r in 0..a.rows() {
                            for c in 0..a.cols() {
                                match a.get(r, c) {
                                    Some(Value::Error(e)) => {
                                        return Some(EvalResult::Error(e.clone()))
                                    }
                                    Some(v) => parts.push(EvalResult::from_value(v).to_text()),
                                    None => parts.push(String::new()),
                                }
                            }
                        }
                    }
                    EvalResult::Error(e) => return Some(EvalResult::Error(e)),
                    EvalResult::Empty => parts.push(String::new()),
                    other => parts.push(other.to_text()),
                }
            }
            EvalResult::Text(parts.join(&delimiter))
        }
        // The query language lives in its own module.
        "QUERY" => super::eval_query::evaluate_query(args, lookup),
        _ => return None,
    };
    Some(result)
}

/// `local@domain.tld`: the local part from the characters an address may use,
/// with no leading, trailing or doubled dot; the domain a host name.
fn is_email(text: &str) -> bool {
    let Some((local, domain)) = text.split_once('@') else {
        return false;
    };
    const SPECIAL: &str = ".!#$%&'*+/=?^_`{|}~-";
    !local.is_empty()
        && local.len() <= 64
        && local
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || SPECIAL.contains(c))
        && !local.starts_with('.')
        && !local.ends_with('.')
        && !local.contains("..")
        && is_host(domain)
}

/// An optional http, https or ftp scheme, a host name (or IPv4 address), an
/// optional port, then an optional path, query or fragment; no spaces.
fn is_url(text: &str) -> bool {
    if text.is_empty() || text.chars().any(char::is_whitespace) {
        return false;
    }
    let rest = match text.split_once("://") {
        Some((scheme, rest)) => {
            if !["http", "https", "ftp"]
                .iter()
                .any(|s| scheme.eq_ignore_ascii_case(s))
            {
                return false;
            }
            rest
        }
        None => text,
    };
    let end = rest.find(['/', '?', '#']).unwrap_or(rest.len());
    let authority = &rest[..end];
    let host = match authority.rsplit_once(':') {
        Some((host, port)) => {
            if port.is_empty() || port.len() > 5 || !port.bytes().all(|b| b.is_ascii_digit()) {
                return false;
            }
            host
        }
        None => authority,
    };
    is_host(host) || is_ipv4(host)
}

/// Dot-separated labels of letters, digits and inner hyphens, ending in a
/// top-level domain of two or more letters.
fn is_host(host: &str) -> bool {
    let labels: Vec<&str> = host.split('.').collect();
    labels.len() >= 2
        && labels.iter().all(|l| {
            !l.is_empty()
                && l.len() <= 63
                && l.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'-')
                && !l.starts_with('-')
                && !l.ends_with('-')
        })
        && labels
            .last()
            .is_some_and(|tld| tld.len() >= 2 && tld.bytes().all(|b| b.is_ascii_alphabetic()))
}

fn is_ipv4(host: &str) -> bool {
    let parts: Vec<&str> = host.split('.').collect();
    parts.len() == 4
        && parts
            .iter()
            .all(|p| !p.is_empty() && p.len() <= 3 && p.parse::<u8>().is_ok())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn email_shapes() {
        for good in [
            "a@b.co",
            "first.last+tag@example.com",
            "x_y@sub.example.org",
        ] {
            assert!(is_email(good), "{good}");
        }
        for bad in [
            "",
            "a@b",
            "@b.com",
            "a@@b.com",
            ".a@b.com",
            "a..b@c.com",
            "a b@c.com",
            "a@b.c",
            "a@-b.com",
        ] {
            assert!(!is_email(bad), "{bad}");
        }
    }

    #[test]
    fn url_shapes() {
        for good in [
            "http://www.google.com",
            "https://example.com/path?q=1#top",
            "www.google.com",
            "google.com",
            "ftp://files.example.org:21/pub",
            "http://192.168.1.1:8080",
        ] {
            assert!(is_url(good), "{good}");
        }
        for bad in [
            "",
            "google",
            "http://",
            "mailto:a@b.com",
            "http://exa mple.com",
            "javascript://x.com",
            "example.c",
            "http://x.com:abc",
        ] {
            assert!(!is_url(bad), "{bad}");
        }
    }
}
