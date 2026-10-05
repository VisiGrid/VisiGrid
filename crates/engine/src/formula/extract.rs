//! Extract literal references without rewriting strings, sheet names, Table
//! headers, or unrelated formula spelling. The AST verifies every source edit.
use super::parser::{parse, Expr, ParsedExpr};
use crate::sheet::UnboundSheetRef;
use std::ops::Range;

fn literal(expr: &ParsedExpr) -> bool {
    matches!(expr, Expr::CellRef { .. } | Expr::Range { .. })
}

fn same(a: &ParsedExpr, b: &ParsedExpr) -> bool {
    fn normalized(mut e: ParsedExpr) -> ParsedExpr {
        if let Expr::CellRef { sheet, .. } | Expr::Range { sheet, .. } = &mut e {
            if let UnboundSheetRef::Named(s) = sheet {
                *s = s.to_ascii_uppercase();
            }
        }
        e
    }
    normalized(a.clone()) == normalized(b.clone())
}

fn quoted(bytes: &[u8], mut i: usize) -> usize {
    let quote = bytes[i];
    i += 1;
    while i < bytes.len() {
        if bytes[i] == quote {
            i += 1;
            if bytes.get(i) == Some(&quote) {
                i += 1;
            } else {
                break;
            }
        } else {
            i += 1;
        }
    }
    i
}
fn space(bytes: &[u8], mut i: usize) -> usize {
    while bytes.get(i).is_some_and(u8::is_ascii_whitespace) {
        i += 1;
    }
    i
}
fn word(bytes: &[u8], mut i: usize) -> usize {
    while bytes
        .get(i)
        .is_some_and(|b| b.is_ascii_alphanumeric() || b"_.$".contains(b))
    {
        i += 1;
    }
    i
}

fn spans(source: &str) -> Result<Vec<(Range<usize>, ParsedExpr)>, String> {
    let b = source.as_bytes();
    let mut result = Vec::new();
    let mut i = 0;
    while i < b.len() {
        match b[i] {
            b'"' => i = quoted(b, i),
            b'[' => i += super::structured::bracket_len(&source[i..])?,
            b'0'..=b'9' | b'.' => {
                while b.get(i).is_some_and(|c| c.is_ascii_digit() || *c == b'.') {
                    i += 1;
                }
                if matches!(b.get(i), Some(b'e' | b'E')) {
                    i += 1;
                    if matches!(b.get(i), Some(b'+' | b'-')) {
                        i += 1;
                    }
                    while b.get(i).is_some_and(u8::is_ascii_digit) {
                        i += 1;
                    }
                }
            }
            b'\'' | b'a'..=b'z' | b'A'..=b'Z' | b'_' | b'$' => {
                let start = i;
                let was_quoted = b[i] == b'\'';
                i = if was_quoted { quoted(b, i) } else { word(b, i) };
                let next = space(b, i);
                if b.get(next) == Some(&b'!') {
                    i = word(b, space(b, next + 1));
                } else if was_quoted || matches!(b.get(next), Some(b'(' | b'[')) {
                    continue;
                }
                let next = space(b, i);
                if b.get(next) == Some(&b':') {
                    i = word(b, space(b, next + 1));
                }
                if let Ok(expr) = parse(&format!("={}", &source[start..i])) {
                    if literal(&expr) {
                        result.push((start..i, expr));
                    }
                }
            }
            _ => i += 1,
        }
    }
    Ok(result)
}

pub fn first_reference_literal(source: &str) -> Result<Option<String>, String> {
    parse(source)?;
    Ok(spans(source)?
        .first()
        .map(|(span, _)| source[span.clone()].into()))
}

pub fn reference_literal_count(source: &str, reference: &str) -> Result<usize, String> {
    let target = parse(&format!("={reference}"))?;
    if !literal(&target) {
        return Err("Choose a cell or bounded range reference.".into());
    }
    let count = spans(source)?
        .iter()
        .filter(|(_, e)| same(e, &target))
        .count();
    if count > 0 {
        parse(source)?;
    }
    Ok(count)
}

fn rewrite_expr(
    expr: &mut ParsedExpr,
    target: &ParsedExpr,
    name: &str,
    shadowed: bool,
) -> Result<usize, String> {
    if literal(expr) && same(expr, target) {
        if shadowed {
            return Err(
                "This name is used by LET or LAMBDA in an affected formula. Choose another name."
                    .into(),
            );
        }
        *expr = Expr::NamedRange(name.to_ascii_uppercase());
        return Ok(1);
    }
    let mut count = 0;
    match expr {
        Expr::Function {
            name: function,
            args,
        } if function == "LET" && args.len() >= 3 && args.len() % 2 == 1 => {
            let mut shadowed = shadowed;
            let last = args.len() - 1;
            for pair in args[..last].chunks_mut(2) {
                count += rewrite_expr(&mut pair[1], target, name, shadowed)?;
                shadowed |= matches!(&pair[0], Expr::NamedRange(n) if n.eq_ignore_ascii_case(name));
            }
            count += rewrite_expr(&mut args[last], target, name, shadowed)?;
        }
        Expr::Function {
            name: function,
            args,
        } if function == "LAMBDA" && !args.is_empty() => {
            let last = args.len() - 1;
            let shadowed = shadowed
                || args[..last]
                    .iter()
                    .any(|e| matches!(e, Expr::NamedRange(n) if n.eq_ignore_ascii_case(name)));
            count += rewrite_expr(&mut args[last], target, name, shadowed)?;
        }
        Expr::Function { args, .. } => {
            for arg in args {
                count += rewrite_expr(arg, target, name, shadowed)?;
            }
        }
        Expr::BinaryOp { left, right, .. } => {
            count += rewrite_expr(left, target, name, shadowed)?;
            count += rewrite_expr(right, target, name, shadowed)?;
        }
        _ => {}
    }
    Ok(count)
}

pub fn replace_reference_literal(
    source: &str,
    reference: &str,
    name: &str,
) -> Result<(String, usize), String> {
    let target = parse(&format!("={reference}"))?;
    if !literal(&target) {
        return Err("Choose a cell or bounded range reference.".into());
    }
    let hits: Vec<_> = spans(source)?
        .into_iter()
        .filter(|(_, e)| same(e, &target))
        .collect();
    if hits.is_empty() {
        return Ok((source.into(), 0));
    }
    let mut expected = parse(source)?;
    let count = rewrite_expr(&mut expected, &target, name, false)?;
    let mut result = String::new();
    let mut last = 0;
    for (span, _) in &hits {
        result.push_str(&source[last..span.start]);
        result.push_str(name);
        last = span.end;
    }
    result.push_str(&source[last..]);
    if count != hits.len() || parse(&result).ok().as_ref() != Some(&expected) {
        return Err(
            "A reference cannot be extracted safely from this formula. Nothing was changed.".into(),
        );
    }
    Ok((result, count))
}
