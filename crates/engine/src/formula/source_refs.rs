//! Reference islands in otherwise unsupported formula syntax. Preserve all
//! surrounding bytes, strings and structured headers. Ambiguous reference
//! syntax is refused rather than guessing at which cells it names.
use super::parser::{self, Expr, ParsedExpr};
use std::ops::Range;

pub(crate) fn references(source: &str) -> Result<Vec<(Range<usize>, ParsedExpr)>, String> {
    fn quoted(s: &str, mut i: usize, quote: u8) -> Result<usize, String> {
        i += 1;
        while let Some(&b) = s.as_bytes().get(i) {
            i += 1;
            if b == quote {
                if s.as_bytes().get(i) == Some(&quote) {
                    i += 1;
                } else {
                    return Ok(i);
                }
            }
        }
        Err("An unterminated formula literal cannot be rewritten safely.".into())
    }
    fn atom(s: &str, mut i: usize) -> Result<usize, String> {
        if s.as_bytes().get(i) == Some(&b'\'') {
            return quoted(s, i, b'\'');
        }
        while let Some(c) = s[i..].chars().next() {
            if !c.is_alphanumeric() && !"_.$".contains(c) {
                break;
            }
            i += c.len_utf8();
        }
        Ok(i)
    }
    fn skip_space(s: &str, i: usize) -> usize {
        s.len() - s[i..].trim_start().len()
    }
    fn parse_reference(source: &str) -> Result<ParsedExpr, String> {
        if let Ok(expr @ (Expr::CellRef { .. } | Expr::Range { .. } | Expr::WholeRange { .. })) =
            parser::parse(&format!("={source}"))
        {
            return Ok(expr);
        }
        // Excel can repeat the qualifier at a range's second endpoint. The
        // engine parser accepts its equivalent single-qualifier spelling.
        if let Some((left, right)) = source.split_once(':') {
            if let (
                Ok(Expr::CellRef {
                    sheet,
                    row,
                    col,
                    row_abs,
                    col_abs,
                }),
                Ok(Expr::CellRef {
                    sheet: end_sheet,
                    row: end_row,
                    col: end_col,
                    row_abs: end_row_abs,
                    col_abs: end_col_abs,
                }),
            ) = (
                parser::parse(&format!("={left}")),
                parser::parse(&format!("={right}")),
            ) {
                if sheet == end_sheet {
                    return Ok(Expr::Range {
                        sheet,
                        start_row: row,
                        start_col: col,
                        end_row,
                        end_col,
                        start_row_abs: row_abs,
                        start_col_abs: col_abs,
                        end_row_abs,
                        end_col_abs,
                    });
                }
            }
        }
        Err("A reference cannot be rewritten safely.".into())
    }
    let mut result = Vec::new();
    let mut i = 0;
    while i < source.len() {
        let start = i;
        if source.as_bytes()[i] == b'#' {
            if let Some(error) = [
                "#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#NUM!", "#N/A", "#NULL!", "#SPILL!",
                "#CYCLE!", "#ERROR!",
            ]
            .iter()
            .find(|error| source[i..].starts_with(**error))
            {
                i += error.len();
                continue;
            }
        }
        if source.as_bytes()[i] == b'"' {
            i = quoted(source, i, b'"')?;
            continue;
        }
        if source.as_bytes()[i] == b'[' {
            i += super::structured::bracket_len(&source[i..])?;
            // External workbook qualifiers cannot be interpreted as local ones.
            if source[i..]
                .chars()
                .next()
                .is_some_and(|c| c.is_alphanumeric())
            {
                let name_end = atom(source, i)?;
                let bang = skip_space(source, name_end);
                if source.as_bytes().get(bang) != Some(&b'!') {
                    return Err("An external reference cannot be rewritten safely.".into());
                }
                let reference = skip_space(source, bang + 1);
                i = atom(source, reference)?;
                let next = skip_space(source, i);
                if source.as_bytes().get(next) == Some(&b':') {
                    i = atom(source, skip_space(source, next + 1))?;
                }
                if source.as_bytes().get(skip_space(source, i)) == Some(&b'!') {
                    return Err("An external range cannot be rewritten safely.".into());
                }
                let qualified = format!(
                    "'{}'!{}",
                    source[start..name_end].replace('\'', "''"),
                    &source[reference..i]
                );
                result.push((start..i, parse_reference(&qualified)?));
            }
            continue;
        }
        i = atom(source, i)?;
        if i == start {
            i += source[i..].chars().next().unwrap().len_utf8();
            continue;
        }
        let mut next = skip_space(source, i);
        if source.as_bytes().get(next) == Some(&b'[') {
            i = next + super::structured::bracket_len(&source[next..])?;
            if let Ok(expr @ Expr::StructuredRef(_)) =
                parser::parse(&format!("={}", &source[start..i]))
            {
                result.push((start..i, expr));
            }
            continue;
        }
        let qualified = source.as_bytes().get(next) == Some(&b'!');
        if qualified {
            let reference = skip_space(source, next + 1);
            i = atom(source, reference)?;
            if i == reference {
                return Err("A sheet-qualified reference cannot be rewritten safely.".into());
            }
            next = skip_space(source, i);
        }
        let range = source.as_bytes().get(next) == Some(&b':');
        if range {
            let endpoint = skip_space(source, next + 1);
            i = atom(source, endpoint)?;
            if i == endpoint {
                return Err("A range reference cannot be rewritten safely.".into());
            }
            next = skip_space(source, i);
            if source.as_bytes().get(next) == Some(&b'!') {
                let endpoint = skip_space(source, next + 1);
                i = atom(source, endpoint)?;
                if i == endpoint {
                    return Err("A qualified range cannot be rewritten safely.".into());
                }
                next = skip_space(source, i);
            }
        }
        if source.as_bytes().get(next) == Some(&b'(') && !qualified {
            continue;
        }
        match parse_reference(&source[start..i]) {
            Ok(expr) => result.push((start..i, expr)),
            _ if qualified || range => {
                return Err("A sheet-qualified reference cannot be rewritten safely.".into())
            }
            _ => {}
        }
    }
    Ok(result)
}

pub(crate) fn rewrite(
    source: &str,
    mut adjust: impl FnMut(&mut ParsedExpr) -> bool,
) -> Result<String, String> {
    let mut result = source.to_owned();
    for (span, mut expr) in references(source)?.into_iter().rev() {
        if adjust(&mut expr) {
            let text = parser::format_parsed_expr(&expr);
            result.replace_range(span, text.trim_start_matches('='));
        }
    }
    Ok(result)
}
