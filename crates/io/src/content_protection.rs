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
    let text = n.to_string();
    let (negative, unsigned) = match text.strip_prefix('-') {
        Some(rest) => (true, rest),
        None => (false, text.as_str()),
    };
    let (mantissa, exponent) = unsigned
        .split_once(['e', 'E'])
        .map(|(m, e)| (m, e.parse::<i32>().expect("JSON number exponent")))
        .unwrap_or((unsigned, 0));
    let fraction = mantissa.split_once('.').map_or(0, |(_, f)| f.len() as i32);
    let digits = mantissa.replace('.', "");
    let significant = digits.trim_start_matches('0');
    if significant.is_empty() {
        return (false, "0".into(), 0);
    }
    let normalized = significant.trim_end_matches('0');
    let trailing = significant.len() - normalized.len();
    (
        negative,
        normalized.into(),
        exponent - fraction + trailing as i32,
    )
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
}
