//! Rename parsed sheet qualifiers while preserving the rest of each source.
use super::parser::{parse, Expr, ParsedExpr};
use crate::sheet::{normalize_sheet_name, UnboundSheetRef};

fn rename_expr(expr: &mut ParsedExpr, old: &str, new: &str) -> usize {
    match expr {
        Expr::CellRef { sheet, .. }
        | Expr::Range { sheet, .. }
        | Expr::WholeRange { sheet, .. } => {
            if let UnboundSheetRef::Named(name) = sheet {
                if normalize_sheet_name(name) == normalize_sheet_name(old) {
                    *name = new.into();
                    return 1;
                }
            }
            0
        }
        Expr::Function { args, .. } => args.iter_mut().map(|e| rename_expr(e, old, new)).sum(),
        Expr::BinaryOp { left, right, .. } => {
            rename_expr(left, old, new) + rename_expr(right, old, new)
        }
        _ => 0,
    }
}

pub fn rename_sheet_reference(source: &str, old: &str, new: &str) -> Result<String, String> {
    let mut spans = Vec::new();
    let b = source.as_bytes();
    let mut i = 0;
    while i < b.len() {
        let start = i;
        let name = match b[i] {
            b'"' | b'\'' => {
                let quote = b[i];
                i += 1;
                while i < b.len() {
                    if b[i] == quote {
                        i += 1;
                        if b.get(i) == Some(&quote) {
                            i += 1;
                        } else {
                            break;
                        }
                    } else {
                        i += 1;
                    }
                }
                if quote == b'"' {
                    continue;
                }
                source[start..i]
                    .strip_prefix('\'')
                    .and_then(|s| s.strip_suffix('\''))
                    .map(|s| s.replace("''", "'"))
            }
            b'[' => {
                i += super::structured::bracket_len(&source[i..])?;
                continue;
            }
            _ => {
                while i < b.len() {
                    let c = source[i..].chars().next().unwrap();
                    if !c.is_alphanumeric() && !"_.$".contains(c) {
                        break;
                    }
                    i += c.len_utf8();
                }
                if start == i {
                    i += source[i..].chars().next().unwrap().len_utf8();
                    continue;
                }
                Some(source[start..i].into())
            }
        };
        if source[i..].trim_start().starts_with('!')
            && name.is_some_and(|n| normalize_sheet_name(&n) == normalize_sheet_name(old))
        {
            spans.push(start..i);
        }
    }
    if spans.is_empty() {
        return Ok(source.into());
    }
    let input = |s: &str| {
        if s.starts_with('=') {
            s.into()
        } else {
            format!("={s}")
        }
    };
    let mut expected = parse(&input(source)).map_err(|_| {
        "A formula referencing this sheet cannot be parsed. Resolve it before renaming the sheet."
    })?;
    let count = rename_expr(&mut expected, old, new);
    // Quoting every rewritten qualifier also handles Unicode, punctuation,
    // numeric names and embedded apostrophes without changing their meaning.
    let qualifier = format!("'{}'", new.replace('\'', "''"));
    let mut result = String::new();
    let mut last = 0;
    for span in &spans {
        result.push_str(&source[last..span.start]);
        result.push_str(&qualifier);
        last = span.end;
    }
    result.push_str(&source[last..]);
    if count != spans.len() || parse(&input(&result)).ok().as_ref() != Some(&expected) {
        return Err("A sheet reference cannot be rewritten safely. Nothing was renamed.".into());
    }
    Ok(result)
}
