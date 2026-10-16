//! Formulas in a Google Sheets .xlsx export.
//!
//! Google writes a formula Excel lacks (FILTER on old Excel, ARRAYFORMULA,
//! REGEXMATCH, SPLIT, QUERY...) as
//! `=IFERROR(__xludf.DUMMYFUNCTION("<the Sheets formula>"),<last value>)`, and
//! each other cell of an array result as
//! `=IFERROR(__xludf.DUMMYFUNCTION("""COMPUTED_VALUE"""),<value>)`. Excel shows
//! the frozen last value and never recalculates it. The original formula is
//! intact inside the string, so import restores it: what the engine supports
//! calculates again, and what it doesn't is counted as an unsupported function
//! in the import report instead of hiding behind an old value.

const DUMMY: &str = "__xludf.DUMMYFUNCTION(";
const COMPUTED_VALUE: &str = "\"COMPUTED_VALUE\"";

#[derive(Debug, PartialEq)]
pub(crate) enum GoogleFormula {
    /// Not a Google wrapper: import the formula as it is.
    NotWrapped,
    /// The Sheets formula, with its leading `=`.
    Restored(String),
    /// A cell an array result spilled into: it imports empty, so the restored
    /// array formula can spill there again.
    SpillCell,
}

/// Recognises the wrapper only when it is the whole formula (optionally inside
/// Google's IFERROR), which is the only way Google writes it.
pub(crate) fn unwrap(formula: &str) -> GoogleFormula {
    let body = formula.strip_prefix('=').unwrap_or(formula).trim();
    let (inner, rest) = match strip_prefix_ci(body, "IFERROR(") {
        Some(after) => match strip_prefix_ci(after.trim_start(), DUMMY) {
            Some(after) => (after, true),
            None => return GoogleFormula::NotWrapped,
        },
        None => match strip_prefix_ci(body, DUMMY) {
            Some(after) => (after, false),
            None => return GoogleFormula::NotWrapped,
        },
    };
    let Some((text, after)) = string_literal(inner.trim_start()) else {
        return GoogleFormula::NotWrapped;
    };
    let Some(after) = after.trim_start().strip_prefix(')') else {
        return GoogleFormula::NotWrapped;
    };
    let after = after.trim_start();
    // The IFERROR's second argument is the cached value; DUMMYFUNCTION alone
    // must end the formula.
    let closed = if rest {
        after.strip_prefix(',').is_some_and(closes_at_end)
    } else {
        after.is_empty()
    };
    if !closed {
        return GoogleFormula::NotWrapped;
    }
    if text == COMPUTED_VALUE {
        return GoogleFormula::SpillCell;
    }
    let text = text.trim();
    if text.is_empty() {
        return GoogleFormula::NotWrapped;
    }
    GoogleFormula::Restored(if text.starts_with('=') {
        text.to_owned()
    } else {
        format!("={text}")
    })
}

/// Whether the first unmatched `)` in `s`, outside strings and quoted sheet
/// names, is its last character.
fn closes_at_end(s: &str) -> bool {
    let mut depth = 0usize;
    let mut quote = None;
    for (i, ch) in s.char_indices() {
        match quote {
            // A doubled delimiter closes and reopens: same result.
            Some(q) if ch == q => quote = None,
            Some(_) => {}
            None => match ch {
                '"' | '\'' => quote = Some(ch),
                '(' | '{' => depth += 1,
                '}' => depth = depth.saturating_sub(1),
                ')' if depth == 0 => return i + 1 == s.len(),
                ')' => depth -= 1,
                _ => {}
            },
        }
    }
    false
}

fn strip_prefix_ci<'a>(s: &'a str, prefix: &str) -> Option<&'a str> {
    s.get(..prefix.len())
        .filter(|head| head.eq_ignore_ascii_case(prefix))
        .map(|_| &s[prefix.len()..])
}

/// A formula string literal at the start of `s`: its text with doubled quotes
/// undone, and what follows the closing quote.
fn string_literal(s: &str) -> Option<(String, &str)> {
    let mut rest = s.strip_prefix('"')?;
    let mut text = String::new();
    loop {
        let end = rest.find('"')?;
        text.push_str(&rest[..end]);
        rest = &rest[end + 1..];
        match rest.strip_prefix('"') {
            Some(after) => {
                text.push('"');
                rest = after;
            }
            None => return Some((text, rest)),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn restores_the_sheets_formula_from_googles_wrapper() {
        assert_eq!(
            unwrap(r#"=IFERROR(__xludf.DUMMYFUNCTION("FILTER(A1:A3,B1:B3>4)"),"b")"#),
            GoogleFormula::Restored("=FILTER(A1:A3,B1:B3>4)".into())
        );
        // calamine gives the formula without its `=`; case does not matter.
        assert_eq!(
            unwrap(r#"iferror(__XLUDF.DummyFunction("ARRAYFORMULA(A1:A3*2)"),6)"#),
            GoogleFormula::Restored("=ARRAYFORMULA(A1:A3*2)".into())
        );
        assert_eq!(
            unwrap(r#"=__xludf.DUMMYFUNCTION("SPLIT(A1,"","")")"#),
            GoogleFormula::Restored(r#"=SPLIT(A1,",")"#.into())
        );
    }

    #[test]
    fn quotes_inside_the_formula_survive() {
        assert_eq!(
            unwrap(r#"=IFERROR(__xludf.DUMMYFUNCTION("REGEXMATCH(A1,""^a"")"),TRUE)"#),
            GoogleFormula::Restored(r#"=REGEXMATCH(A1,"^a")"#.into())
        );
        assert_eq!(
            unwrap(
                r#"=IFERROR(__xludf.DUMMYFUNCTION("QUERY(A1:B9,""select A where B > 4"",1)"),"x, y")"#
            ),
            GoogleFormula::Restored(r#"=QUERY(A1:B9,"select A where B > 4",1)"#.into())
        );
        // A cached value with parentheses and quotes of its own.
        assert_eq!(
            unwrap(r#"=IFERROR(__xludf.DUMMYFUNCTION("UNIQUE(A1:A3)"),"(a)"")")"#),
            GoogleFormula::Restored("=UNIQUE(A1:A3)".into())
        );
    }

    #[test]
    fn array_overflow_cells_import_empty() {
        assert_eq!(
            unwrap(r#"=IFERROR(__xludf.DUMMYFUNCTION("""COMPUTED_VALUE"""),"c")"#),
            GoogleFormula::SpillCell
        );
    }

    #[test]
    fn other_formulas_are_left_alone() {
        for f in [
            "=SUM(A1:A3)",
            r#"=IFERROR(A1/B1,"")"#,
            // The wrapper inside a larger formula is not how Google writes it.
            r#"=1+__xludf.DUMMYFUNCTION("A1")"#,
            r#"=IFERROR(__xludf.DUMMYFUNCTION("A1"),1)+1"#,
            r#"=IFERROR(__xludf.DUMMYFUNCTION("A1"),1)+SUM(B1)"#,
            r#"=LEN("__xludf.DUMMYFUNCTION(")"#,
            r#"=__xludf.DUMMYFUNCTION("")"#,
            r#"=__xludf.DUMMYFUNCTION("unterminated)"#,
        ] {
            assert_eq!(unwrap(f), GoogleFormula::NotWrapped, "{f}");
        }
    }
}
