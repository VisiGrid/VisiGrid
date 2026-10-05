//! Optional, versioned validation metadata. Older files omit the key.
use rusqlite::{params, Connection, OptionalExtension};
use serde::{Deserialize, Serialize};
use visigrid_engine::{
    sheet::Sheet,
    validation::{CellRange, ValidationRule, ValidationStore},
};

const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_RANGES: usize = 100_000;

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Saved {
    version: u32,
    rules: Vec<(CellRange, ValidationRule)>,
    exclusions: Vec<CellRange>,
}

impl Saved {
    fn validate(&self, sheet: &Sheet) -> Result<(), String> {
        if self.version != 1 {
            return Err("Validation metadata has an invalid version.".into());
        }
        if self.rules.len().saturating_add(self.exclusions.len()) > MAX_RANGES {
            return Err("Validation metadata exceeds the supported range limit.".into());
        }
        let valid = |r: &CellRange| {
            r.start_row <= r.end_row
                && r.start_col <= r.end_col
                && r.end_row < sheet.rows
                && r.end_col < sheet.cols
        };
        if !self.rules.iter().all(|(r, _)| valid(r)) || !self.exclusions.iter().all(valid) {
            return Err("Validation metadata contains a corrupt worksheet range.".into());
        }
        let mut ranges = std::collections::HashSet::new();
        if self.rules.iter().any(|(r, _)| !ranges.insert(*r)) {
            return Err("Validation metadata contains duplicate rule ranges.".into());
        }
        Ok(())
    }
}

pub(super) fn save(conn: &Connection, index: usize, sheet: &Sheet) -> Result<(), String> {
    if sheet.validations.is_empty() && sheet.validations.exclusions_len() == 0 {
        return Ok(());
    }
    let mut saved = Saved {
        version: 1,
        rules: sheet
            .validations
            .iter()
            .map(|(r, v)| (*r, v.clone()))
            .collect(),
        exclusions: sheet.validations.exclusions_iter().copied().collect(),
    };
    let key = |r: &CellRange| (r.start_row, r.start_col, r.end_row, r.end_col);
    saved.rules.sort_by_key(|(r, _)| key(r));
    saved.exclusions.sort_by_key(key);
    saved.validate(sheet)?;
    let raw = serde_json::to_string(&saved).map_err(|e| e.to_string())?;
    if raw.len() > MAX_BYTES {
        return Err("Validation metadata exceeds 8 MiB.".into());
    }
    conn.execute(
        "INSERT OR REPLACE INTO meta (key,value) VALUES (?1,?2)",
        params![format!("validations_{index}"), raw],
    )
    .map_err(|e| e.to_string())?;
    Ok(())
}

pub(super) fn load(conn: &Connection, index: usize, sheet: &mut Sheet) -> Option<String> {
    let read = || -> Result<Option<ValidationStore>, String> {
        let key = format!("validations_{index}");
        let size: Option<usize> = conn
            .query_row(
                "SELECT length(CAST(value AS BLOB)) FROM meta WHERE key = ?1",
                [&key],
                |r| r.get(0),
            )
            .optional()
            .map_err(|e| e.to_string())?;
        let Some(size) = size else {
            return Ok(None);
        };
        if size > MAX_BYTES {
            return Err("Validation metadata exceeds 8 MiB.".into());
        }
        let raw: String = conn
            .query_row("SELECT value FROM meta WHERE key = ?1", [&key], |r| {
                r.get(0)
            })
            .map_err(|e| e.to_string())?;
        #[derive(Deserialize)]
        struct Version {
            version: u32,
        }
        let version: Version = serde_json::from_str(&raw)
            .map_err(|e| format!("Validation metadata is corrupt: {e}"))?;
        if version.version > 1 {
            return Err(
                "These validation rules need a newer VisiGrid. Upgrade to edit this file.".into(),
            );
        }
        let saved: Saved = serde_json::from_str(&raw)
            .map_err(|e| format!("Validation metadata is corrupt: {e}"))?;
        saved.validate(sheet)?;
        let mut store = ValidationStore::new();
        for (range, rule) in saved.rules {
            store.set(range, rule);
        }
        for range in saved.exclusions {
            store.exclude(range);
        }
        Ok(Some(store))
    };
    match read() {
        Ok(Some(store)) => {
            sheet.validations = store;
            None
        }
        Ok(None) => None,
        Err(error) => Some(format!(
            "{error} Validation rules were not restored; the cells are available read-only."
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::native;
    use visigrid_engine::{
        validation::{ListSource, ValidationType},
        workbook::Workbook,
    };

    #[test]
    fn every_native_writer_retains_rules_and_exclusions() {
        let mut wb = Workbook::new();
        let mut rule = ValidationRule::list_range("Choices");
        rule.rule_type = ValidationType::List(ListSource::NamedRange("Choices".into()));
        let range = CellRange::new(1, 1, 8, 1);
        wb.active_sheet_mut().validations.set(range, rule.clone());
        wb.active_sheet_mut()
            .validations
            .exclude(CellRange::new(3, 1, 3, 1));
        wb.define_name_for_range("Choices", 0, 10, 0, 11, 0)
            .unwrap();
        for writer in 0..4 {
            let dir = tempfile::tempdir().unwrap();
            let path = dir.path().join("validation.sheet");
            match writer {
                0 => native::save(wb.active_sheet(), &path).unwrap(),
                1 => native::save_workbook(&wb, &path).unwrap(),
                2 => native::save_workbook_with_metadata(&wb, &Default::default(), &path).unwrap(),
                _ => native::save_workbook_full(&wb, &Default::default(), &[], &[], &path).unwrap(),
            }
            let loaded = native::load_workbook(&path).unwrap();
            assert!(loaded.read_only_reason().is_none());
            assert_eq!(loaded.active_sheet().validations.get(2, 1), Some(&rule));
            assert!(loaded.active_sheet().validations.is_excluded(3, 1));
        }
    }

    #[test]
    fn missing_metadata_is_compatible_and_corrupt_or_future_rules_are_read_only() {
        let wb = Workbook::new();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("validation.sheet");
        native::save_workbook(&wb, &path).unwrap();
        let conn = Connection::open(&path).unwrap();
        let key = "validations_0";
        assert_eq!(
            conn.query_row("SELECT count(*) FROM meta WHERE key = ?1", [key], |r| r
                .get::<_, usize>(
                0
            ))
            .unwrap(),
            0
        );
        assert!(native::load_workbook(&path)
            .unwrap()
            .read_only_reason()
            .is_none());
        for raw in [
            "{broken".to_string(),
            r#"{"version":2,"rules":[],"exclusions":[],"future_field":true}"#.into(),
            format!(
                r#"{{"version":1,"rules":[],"exclusions":[{{"start_row":0,"start_col":0,"end_row":{},"end_col":0}}]}}"#,
                wb.active_sheet().rows
            ),
            " ".repeat(MAX_BYTES + 1),
        ] {
            conn.execute(
                "INSERT OR REPLACE INTO meta (key,value) VALUES (?1,?2)",
                params![key, raw],
            )
            .unwrap();
            let loaded = native::load_workbook(&path).unwrap();
            assert!(loaded.read_only_reason().is_some());
            assert!(native::save_workbook(&loaded, &dir.path().join("copy.sheet")).is_err());
        }
    }
}
