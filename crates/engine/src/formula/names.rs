//! Named reference refactoring preserves source spelling and punctuation.
use super::parser::{self, Expr, ParsedExpr};

fn rename_expr(expr: &mut ParsedExpr, old: &str, new: &str) -> bool {
    match expr {
        Expr::NamedRange(name) if name.eq_ignore_ascii_case(old) => {
            *name = new.to_uppercase();
            true
        }
        Expr::Function { args, .. } => args
            .iter_mut()
            .fold(false, |changed, arg| rename_expr(arg, old, new) | changed),
        Expr::BinaryOp { left, right, .. } => {
            rename_expr(left, old, new) | rename_expr(right, old, new)
        }
        _ => false,
    }
}

/// Strings (including INDIRECT arguments), sheet names, function names and
/// structured-reference fields never become named-range references.
pub fn references_name(source: &str, name: &str) -> bool {
    let input = if source.starts_with('=') {
        source.to_string()
    } else {
        format!("={source}")
    };
    parser::parse(&input).is_ok_and(|mut expr| rename_expr(&mut expr, name, name))
}

pub fn rename_reference(source: &str, old: &str, new: &str) -> Result<String, String> {
    let input = if source.starts_with('=') {
        source.to_string()
    } else {
        format!("={source}")
    };
    // Unparsed formulas are not silently rewritten. Refuse if they may contain
    // the old name; unrelated imported unsupported formulas remain untouched.
    let mut expected = match parser::parse(&input) {
        Ok(expr) => expr,
        Err(_)
            if source
                .to_ascii_lowercase()
                .contains(&old.to_ascii_lowercase()) =>
        {
            return Err(
                "A formula using this name cannot be parsed. Resolve it before renaming the range."
                    .into(),
            )
        }
        Err(_) => return Ok(source.into()),
    };
    if !rename_expr(&mut expected, old, new) {
        return Ok(source.into());
    }
    let bytes = source.as_bytes();
    let mut i = 0;
    let mut last = 0;
    let mut result = String::new();
    while i < bytes.len() {
        match bytes[i] {
            b'"' | b'\'' => {
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
            }
            b'[' => {
                i += super::structured::bracket_len(&source[i..])?;
            }
            b'0'..=b'9' | b'.' => {
                while i < bytes.len() && (bytes[i].is_ascii_digit() || bytes[i] == b'.') {
                    i += 1;
                }
                if matches!(bytes.get(i), Some(b'e' | b'E')) {
                    i += 1;
                    if matches!(bytes.get(i), Some(b'+' | b'-')) {
                        i += 1;
                    }
                    while i < bytes.len() && bytes[i].is_ascii_digit() {
                        i += 1;
                    }
                }
            }
            b'a'..=b'z' | b'A'..=b'Z' | b'_' | b'$' => {
                let start = i;
                while i < bytes.len()
                    && (bytes[i].is_ascii_alphanumeric() || b"_.$".contains(&bytes[i]))
                {
                    i += 1;
                }
                let token = &source[start..i];
                let next = source[i..].trim_start().as_bytes().first().copied();
                if token.eq_ignore_ascii_case(old)
                    && !matches!(next, Some(b'(' | b'!' | b'[' | b':'))
                    && !source[..start].trim_end().ends_with(':')
                {
                    result.push_str(&source[last..start]);
                    result.push_str(new);
                    last = i;
                }
            }
            _ => {
                i += 1;
            }
        }
    }
    result.push_str(&source[last..]);
    let check = if result.starts_with('=') {
        result.clone()
    } else {
        format!("={result}")
    };
    if parser::parse(&check).ok().as_ref() != Some(&expected) {
        return Err(
            "This formula's named reference cannot be rewritten safely. Nothing was changed."
                .into(),
        );
    }
    Ok(result)
}
