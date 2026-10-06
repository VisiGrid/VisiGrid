//! Preserve the source when a canonical rewrite cannot represent its content.
use serde_json::Value;
use sha2::{Digest, Sha256};
use visigrid_engine::{sheet::Sheet, workbook::Workbook};

/// Report the first source member/value a projected writer would lose.
/// Extra output members are permitted. Formula cached results may legitimately
/// change after supported recalculation; their surrounding metadata may not.
pub fn first_loss(source: &Value, projected: &Value) -> Option<String> {
    fn visit(a: &Value, b: &Value, path: &str) -> Option<String> {
        match (a, b) {
            (Value::Object(a), Value::Object(b)) => {
                for (key, value) in a {
                    if key == "value"
                        && path.rsplit_once("/cells/").is_some_and(|(_, index)| {
                            !index.is_empty() && index.bytes().all(|c| c.is_ascii_digit())
                        })
                        && a.get("formula")
                            .and_then(Value::as_str)
                            .is_some_and(|f| !f.is_empty())
                    {
                        continue;
                    }
                    let next = format!("{path}/{}", key.replace('~', "~0").replace('/', "~1"));
                    let Some(other) = b.get(key) else {
                        // The codec omits empty cell lists. Recognize that one
                        // default only at actual workbook/sheet bodies; unknown
                        // empty fields must still survive unchanged.
                        let body = (path.is_empty() && a.get("format").and_then(Value::as_str) == Some("visigrid-json"))
                            || path.strip_prefix("/sheets/").is_some_and(|index| !index.is_empty() && index.bytes().all(|c| c.is_ascii_digit()));
                        if key == "cells" && body && value.as_array().is_some_and(Vec::is_empty) {
                            continue;
                        }
                        return Some(next);
                    };
                    if let Some(loss) = visit(value, other, &next) {
                        return Some(loss);
                    }
                }
                None
            }
            (Value::Array(a), Value::Array(b)) => {
                let by_cell: std::collections::BTreeMap<_, _> = b
                    .iter()
                    .filter_map(|v| Some(((v.get("row")?.as_u64()?, v.get("col")?.as_u64()?), v)))
                    .collect();
                for (i, value) in a.iter().enumerate() {
                    // Canonical writers sort cells. Match anchors rather than
                    // assuming an input array was already ordered.
                    let other = if value.get("row").is_some() && value.get("col").is_some() {
                        value
                            .get("row")
                            .and_then(Value::as_u64)
                            .zip(value.get("col").and_then(Value::as_u64))
                            .and_then(|key| by_cell.get(&key).copied())
                    } else {
                        b.get(i)
                    };
                    let next = format!("{path}/{i}");
                    let Some(other) = other else {
                        return Some(next);
                    };
                    if let Some(loss) = visit(value, other, &next) {
                        return Some(loss);
                    }
                }
                None
            }
            (Value::Number(a), Value::Number(b)) if decimal_number(a) == decimal_number(b) => None,
            _ if a == b => None,
            _ => Some(path.to_owned()),
        }
    }
    visit(source, projected, "")
}

// Compare the represented decimal values without converting integers to f64:
// 42 and 42.0 are equivalent, but 9007199254740993 must not match a rounded
// 9007199254740992.0. Keep the exponent separate to avoid large allocations.
fn decimal_number(n: &serde_json::Number) -> (bool, String, i32) {
    decimal_literal(&n.to_string()).expect("representable JSON number")
}

fn decimal_literal(text: &str) -> Option<(bool, String, i32)> {
    let (negative, unsigned) = match text.strip_prefix('-') {
        Some(rest) => (true, rest),
        None => (false, text),
    };
    let (mantissa, exponent) = match unsigned.split_once(['e', 'E']) {
        Some((m,e)) => (m,e.parse::<i32>().ok()?),
        None => (unsigned,0),
    };
    let fraction = match mantissa.split_once('.') {
        Some((_,f)) => i32::try_from(f.len()).ok()?,
        None => 0,
    };
    let digits = mantissa.replace('.', "");
    let significant = digits.trim_start_matches('0');
    if significant.is_empty() { return Some((false,"0".into(),0)); }
    let normalized = significant.trim_end_matches('0');
    let trailing = i32::try_from(significant.len() - normalized.len()).ok()?;
    Some((negative, normalized.into(), exponent.checked_sub(fraction)?.checked_add(trailing)?))
}

#[derive(serde::Deserialize)]
struct NumericCell {
    #[serde(default)]
    value: Option<Box<serde_json::value::RawValue>>,
}

#[derive(Default, serde::Deserialize)]
struct NumericBody {
    #[serde(default)]
    cells: Vec<NumericCell>,
}

#[derive(serde::Deserialize)]
struct NumericDocument {
    #[serde(default)]
    cells: Vec<NumericCell>,
    #[serde(default)]
    sheets: Vec<NumericBody>,
}

fn numeric_cell_loss(cell: &NumericCell) -> bool {
    // Both values and caches must be readable before recalculation is allowed.
    let Some(raw) = &cell.value else { return false };
    let literal = raw.get().trim();
    if !literal.starts_with(|c: char| c == '-' || c.is_ascii_digit()) { return false; }
    let projected = literal.parse::<f64>().ok().and_then(serde_json::Number::from_f64);
    let Some(projected) = projected else { return true };
    match (decimal_literal(literal), decimal_literal(&projected.to_string())) {
        (Some(original),Some(projected)) => original != projected,
        _ => true,
    }
}

// The normal Value parser retains u64 integers but rounds larger integers
// and precise fractions. Inspect original tokens, including opaque metadata,
// while skipping quoted strings and their escapes.
fn parsed_number_loss(content: &str) -> bool {
    let bytes = content.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'"' {
            i += 1;
            while i < bytes.len() {
                match bytes[i] {
                    b'\\' => i += 2,
                    b'"' => { i += 1; break; }
                    _ => i += 1,
                }
            }
        } else if bytes[i] == b'-' || bytes[i].is_ascii_digit() {
            let start = i;
            i += 1;
            while i < bytes.len() && (bytes[i].is_ascii_digit() || matches!(bytes[i], b'.' | b'e' | b'E' | b'+' | b'-')) { i += 1; }
            let literal = &content[start..i];
            let parsed = serde_json::from_str::<Value>(literal).ok();
            let Some(Value::Number(parsed)) = parsed else { return true };
            if decimal_literal(literal) != decimal_literal(&parsed.to_string()) { return true; }
        } else {
            i += 1;
        }
    }
    false
}

pub(crate) fn inline_number_loss(content: &str) -> Result<Option<String>, String> {
    if parsed_number_loss(content) { return Ok(Some("a numeric literal".into())); }
    let document: NumericDocument = serde_json::from_str(content).map_err(|e| e.to_string())?;
    if document.sheets.is_empty() {
        Ok(document.cells.iter().position(numeric_cell_loss).map(|i| format!("/cells/{i}/value")))
    } else {
        Ok(document.sheets.iter().enumerate().find_map(|(s,body)| {
            body.cells.iter().position(numeric_cell_loss).map(|i| format!("/sheets/{s}/cells/{i}/value"))
        }))
    }
}

pub(crate) fn band_cell_number_loss(content: &str) -> Result<bool, String> {
    if parsed_number_loss(content) { return Ok(true); }
    let cell: NumericCell = serde_json::from_str(content).map_err(|e| e.to_string())?;
    Ok(numeric_cell_loss(&cell))
}

pub(crate) fn fingerprint(sheet: &Sheet) -> Result<[u8; 32], String> {
    // Value's object representation sorts keys, including engine HashMaps.
    let value = crate::json::protection_projection(sheet)?;
    let bytes = serde_json::to_vec(&value).map_err(|e| e.to_string())?;
    Ok(Sha256::digest(bytes).into())
}

/// A protected document can be copied exactly, never reconstructed from a
/// partial grid. Any mutation, missing/reordered sheet or cleared warning
/// invalidates this permission independently of the UI's read-only state.
pub(crate) fn original_source(wb: &Workbook) -> Result<Option<&str>, String> {
    if wb.pending_bands.is_some() || wb.sheets().iter().any(|s| s.canonical_content_protection.as_ref().is_some_and(|p| p.incomplete_bands)) {
        return Err("Cannot export or save a workbook before all band data is loaded.".into());
    }
    let Some(first) = wb.sheets().first() else {
        return Ok(None);
    };
    let Some(protection) = &first.canonical_content_protection else {
        if wb
            .sheets()
            .iter()
            .any(|s| s.canonical_content_protection.is_some())
        {
            return Err("Cannot save a partial protected workbook.".into());
        }
        return Ok(None);
    };
    let ids: Vec<_> = wb.sheets().iter().map(|s| s.id).collect();
    if ids != protection.sheet_ids {
        return Err("Cannot change protected workbook structure.".into());
    }
    for sheet in wb.sheets() {
        let Some(p) = &sheet.canonical_content_protection else {
            return Err("Protected workbook metadata is missing.".into());
        };
        if p.source != protection.source || fingerprint(sheet)? != p.fingerprint {
            return Err("Cannot save changes to unsupported workbook content. Export the original or upgrade VisiGrid.".into());
        }
    }
    Ok(Some(protection.source.as_str()))
}

#[cfg(test)]
mod tests {
    use super::first_loss;
    use serde_json::json;

    #[test]
    fn cached_results_are_only_exempt_inside_formula_cells() {
        let before = json!({"sheets":[{"cells":[{"row":0,"col":0,"formula":"=1+1","value":1}]}]});
        let after = json!({"sheets":[{"cells":[{"row":0,"col":0,"formula":"=1+1","value":2}]}]});
        assert_eq!(first_loss(&before, &after), None);
        assert!(first_loss(
            &json!({"future":{"formula":"opaque","value":1}}),
            &json!({"future":{"formula":"opaque","value":2}})
        )
        .is_some());
        assert!(first_loss(
            &json!({"cells":[{"row":0,"col":0,"formula":"","value":1}]}),
            &json!({"cells":[{"row":0,"col":0,"formula":"","value":2}]})
        )
        .is_some());
    }

    #[test]
    fn numeric_representation_changes_are_not_content_loss() {
        for (before, after) in [
            ("42", "42.0"),
            ("1000", "1e3"),
            ("0", "-0.0"),
            ("0.125", "1.25e-1"),
        ] {
            assert_eq!(
                first_loss(
                    &serde_json::from_str(before).unwrap(),
                    &serde_json::from_str(after).unwrap()
                ),
                None
            );
        }
        assert!(first_loss(
            &serde_json::from_str("9007199254740993").unwrap(),
            &serde_json::from_str("9007199254740992.0").unwrap()
        )
        .is_some());
    }
    #[test]
    fn empty_known_cell_lists_can_be_omitted_but_extensions_cannot() {
        assert_eq!(first_loss(&json!({"format":"visigrid-json","cells":[],"sheets":[{"name":"A","cells":[]}]}),
            &json!({"format":"visigrid-json","sheets":[{"name":"A"}]})), None);
        assert_eq!(first_loss(&json!({"format":"visigrid-json","future":[]}), &json!({"format":"visigrid-json"})), Some("/future".into()));
        assert_eq!(first_loss(&json!({"row":0,"col":0,"cells":[]}), &json!({"row":0,"col":0})), Some("/cells".into()));
        assert!(first_loss(&json!({"format":"visigrid-json","cells":[{"row":0,"col":0,"value":42}]}), &json!({"format":"visigrid-json"})).is_some());
    }

}
