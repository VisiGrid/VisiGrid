//! Formulas in a Google Sheets .xlsx export.
//!
//! Google writes a formula Excel lacks (FILTER on old Excel, ARRAYFORMULA,
//! REGEXMATCH, SPLIT, QUERY...) as
//! `=IFERROR(__xludf.DUMMYFUNCTION("<the Sheets formula>"),<last value>)`, and
//! each other cell of an array result as
//! `=IFERROR(__xludf.DUMMYFUNCTION("""COMPUTED_VALUE"""),<value>)`. Excel shows
//! the frozen last value and never recalculates it. The original formula is
//! intact inside the string, so import restores it and what the engine
//! supports calculates again. A formula using a function the engine lacks
//! keeps Google's last value (and its array cells theirs) rather than showing
//! an error, gets a note naming the formula, and is listed in the import
//! report: never silently frozen, never a wall of errors.

use std::collections::HashSet;

const DUMMY: &str = "__xludf.DUMMYFUNCTION(";
const COMPUTED_VALUE: &str = "\"COMPUTED_VALUE\"";

#[derive(Debug, PartialEq)]
pub(crate) enum GoogleFormula {
    /// Not a Google wrapper: import the formula as it is.
    NotWrapped,
    /// The Sheets formula, with its leading `=`.
    Restored(String),
    /// A cell an array result spilled into: it imports empty when its array
    /// formula is live, so the array spills there again.
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

/// The array formula a placeholder belongs to, among `anchors` (every
/// restored Google formula on the sheet). Google's array results are
/// rectangles: go up the placeholder's column while the cells are
/// placeholders, then left along that row; the formula is the cell above (the
/// placeholder was in the formula's column) or to the left (the top row).
pub(crate) fn spill_owner(
    (row, col): (usize, usize),
    placeholders: &HashSet<(usize, usize)>,
    anchors: &HashSet<(usize, usize)>,
) -> Option<(usize, usize)> {
    let mut r = row;
    while r > 0 && placeholders.contains(&(r - 1, col)) {
        r -= 1;
    }
    let left_is_placeholder = col > 0 && placeholders.contains(&(r, col - 1));
    if r > 0 && !left_is_placeholder && anchors.contains(&(r - 1, col)) {
        return Some((r - 1, col));
    }
    let mut c = col;
    while c > 0 && placeholders.contains(&(r, c - 1)) {
        c -= 1;
    }
    (c > 0 && anchors.contains(&(r, c - 1))).then(|| (r, c - 1))
}

/// Notes on the cells that kept Google's value, after the file's own comments
/// are applied (an existing note keeps its text, with ours after it).
pub(crate) fn note_frozen(
    frozen: &[(usize, usize, usize, String)],
    workbook: &mut visigrid_engine::workbook::Workbook,
) {
    use visigrid_engine::cell::CellComment;
    for (sheet, row, col, formula) in frozen {
        let Some(sheet) = workbook.sheet_mut(*sheet) else {
            continue;
        };
        let note = format!(
            "Google Sheets formula, not supported in VisiGrid yet: {formula}\nThis cell shows the value Google Sheets last calculated."
        );
        let text = match sheet.comment(*row, *col) {
            Some(existing) if !existing.text.is_empty() => format!("{}\n\n{note}", existing.text),
            _ => note,
        };
        sheet.set_comment(
            *row,
            *col,
            Some(CellComment {
                text,
                author: "VisiGrid import".into(),
            }),
        );
    }
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
    fn placeholders_find_their_array_formula() {
        // A 3x2 array at B2 (B2:C4) beside a 2x1 array at E2 (E2:E3).
        let anchors: HashSet<_> = [(1, 1), (1, 4)].into();
        let placeholders: HashSet<_> = [(1, 2), (2, 1), (2, 2), (3, 1), (3, 2), (2, 4)].into();
        for cell in [(1, 2), (2, 1), (2, 2), (3, 1), (3, 2)] {
            assert_eq!(
                spill_owner(cell, &placeholders, &anchors),
                Some((1, 1)),
                "{cell:?}"
            );
        }
        assert_eq!(spill_owner((2, 4), &placeholders, &anchors), Some((1, 4)));
        // A placeholder with no formula to find keeps its value.
        let stray: HashSet<_> = [(7, 7)].into();
        assert_eq!(spill_owner((7, 7), &stray, &anchors), None);
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
