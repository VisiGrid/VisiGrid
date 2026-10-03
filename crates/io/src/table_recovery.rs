//! Explicit recovery of cells without interpreting unsupported Table metadata.
use visigrid_engine::workbook::{SavedTableCatalog, Workbook};

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum TableLoadIssue {
    FutureVersion(u64),
    Corrupt(String),
}
impl std::fmt::Display for TableLoadIssue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::FutureVersion(version) => write!(f, "This file uses Table format {version}. Upgrade VisiGrid to open it with working Tables."),
            Self::Corrupt(detail) => write!(f, "Table metadata is corrupt: {detail}"),
        }
    }
}
pub(crate) fn decode_catalog(raw: &str) -> Result<SavedTableCatalog, TableLoadIssue> {
    let value: serde_json::Value =
        serde_json::from_str(raw).map_err(|e| TableLoadIssue::Corrupt(e.to_string()))?;
    match value.get("version").and_then(|v| v.as_u64()) {
        Some(version) if version > 4 => return Err(TableLoadIssue::FutureVersion(version)),
        Some(1..=4) => {}
        _ => {
            return Err(TableLoadIssue::Corrupt(
                "invalid or missing Table format version".into(),
            ))
        }
    }
    serde_json::from_value(value).map_err(|e| TableLoadIssue::Corrupt(e.to_string()))
}

/// Preserve stored results for inspection. These are explicitly stale cached
/// values, not recalculated answers. Missing caches remain unavailable.
pub(crate) fn finish_recovery(
    wb: &mut Workbook,
    issue: &TableLoadIssue,
    cached: &[(usize, usize, usize, crate::CachedFormulaValue)],
) {
    for (sheet_index, row, col, value) in cached {
        if let Some(sheet) = wb.sheet(*sheet_index) {
            use visigrid_engine::formula::eval::Value;
            sheet.cache_computed(
                *row,
                *col,
                match value {
                    crate::CachedFormulaValue::Number(n) => Value::Number(*n),
                    crate::CachedFormulaValue::Text(t) => Value::Text(t.clone()),
                },
            );
        }
    }
    for index in 0..wb.sheet_count() {
        wb.sheet_mut(index).unwrap().read_only_reason = Some(issue.to_string());
    }
    wb.set_auto_recalc(false);
}
