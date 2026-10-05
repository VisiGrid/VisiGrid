//! Workbook names are semantic metadata: restore them before recalculation and
//! before the Table catalog checks its shared namespace.
use std::collections::HashSet;
use visigrid_engine::{
    named_range::{is_valid_name, NamedRange, NamedRangeTarget},
    sheet::{NUM_COLS, NUM_ROWS},
    workbook::Workbook,
};

fn validate(wb: &Workbook, names: &[NamedRange]) -> Result<(), String> {
    if names.len() > 8192 {
        return Err("Named ranges exceed the 8,192 definition limit".into());
    }
    let mut seen = HashSet::new();
    let mut bytes = 0usize;
    for name in names {
        is_valid_name(&name.name)?;
        if !seen.insert(name.name.to_lowercase()) {
            return Err(format!("Duplicate named range: {}", name.name));
        }
        bytes = bytes
            .saturating_add(name.name.len())
            .saturating_add(name.description.as_ref().map_or(0, String::len));
        if bytes > 8 * 1024 * 1024 {
            return Err("Named ranges exceed the 8 MiB text limit".into());
        }
        let (sheet, r0, c0, r1, c1) = match name.target {
            NamedRangeTarget::Cell { sheet, row, col } => (sheet, row, col, row, col),
            NamedRangeTarget::Range {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
            } => (sheet, start_row, start_col, end_row, end_col),
        };
        if sheet >= wb.sheet_count() || r0 > r1 || c0 > c1 || r1 >= NUM_ROWS || c1 >= NUM_COLS {
            return Err(format!("Invalid named range target: {}", name.name));
        }
        if wb
            .tables()
            .any(|(_, table)| table.name.eq_ignore_ascii_case(&name.name))
        {
            return Err(format!("Named range conflicts with Table: {}", name.name));
        }
    }
    Ok(())
}

pub(super) fn export(wb: &Workbook) -> Result<Vec<NamedRange>, String> {
    let mut names: Vec<_> = wb.list_named_ranges().into_iter().cloned().collect();
    names.sort_by(|a, b| a.name.cmp(&b.name));
    validate(wb, &names)?;
    Ok(names)
}

pub(super) fn restore(wb: &mut Workbook, names: &[NamedRange]) -> Result<(), String> {
    validate(wb, names)?;
    for name in names {
        wb.named_ranges_mut().set(name.clone())?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::json::{export_workbook, import_any};

    #[test]
    fn names_survive_json_with_descriptions_cross_sheet_formulas_and_stable_order() {
        let mut wb = Workbook::new();
        let source = wb.add_sheet_named("Options").unwrap();
        wb.set_cell_value_tracked(source, 1, 0, "7");
        wb.set_cell_value_tracked(source, 2, 0, "8");
        wb.named_ranges_mut()
            .set(
                NamedRange::range("Values", source, 1, 0, 2, 0)
                    .with_description("Source choices & totals"),
            )
            .unwrap();
        wb.define_name_for_cell("First", source, 1, 0).unwrap();
        wb.set_cell_value_tracked(0, 0, 0, "=SUM(Values)+First");
        let json = export_workbook(&wb, &[], 0).unwrap();
        assert_eq!(
            serde_json::from_str::<serde_json::Value>(&json).unwrap()["version"],
            5
        );
        let loaded = import_any(&json).unwrap().0;
        assert_eq!(loaded.sheet(0).unwrap().get_display(0, 0), "22");
        assert_eq!(
            loaded.get_named_range("values"),
            wb.get_named_range("Values")
        );
        assert_eq!(export_workbook(&loaded, &[], 0).unwrap(), json);
        let plain = export_workbook(&Workbook::new(), &[], 0).unwrap();
        let plain: serde_json::Value = serde_json::from_str(&plain).unwrap();
        assert_eq!(plain["version"], 2);
        assert!(plain.get("named_ranges").is_none());
    }

    #[test]
    fn invalid_json_names_are_rejected_before_publication() {
        let mut wb = Workbook::new();
        wb.define_name_for_cell("Choice", 0, 0, 0).unwrap();
        let json: serde_json::Value =
            serde_json::from_str(&export_workbook(&wb, &[], 0).unwrap()).unwrap();
        let mut older = json.clone();
        older["version"] = 4.into();
        assert!(import_any(&older.to_string())
            .unwrap_err()
            .contains("require visigrid-json v5"));
        let mut duplicate = json.clone();
        let mut second = duplicate["named_ranges"][0].clone();
        second["name"] = "cHoIcE".into();
        duplicate["named_ranges"]
            .as_array_mut()
            .unwrap()
            .push(second);
        assert!(import_any(&duplicate.to_string())
            .unwrap_err()
            .contains("Duplicate named range"));
        for target in [
            NamedRangeTarget::Cell {
                sheet: 1,
                row: 0,
                col: 0,
            },
            NamedRangeTarget::Cell {
                sheet: 0,
                row: NUM_ROWS,
                col: 0,
            },
            NamedRangeTarget::Cell {
                sheet: 0,
                row: 0,
                col: NUM_COLS,
            },
            NamedRangeTarget::Range {
                sheet: 0,
                start_row: 2,
                start_col: 0,
                end_row: 1,
                end_col: 0,
            },
        ] {
            let mut broken = json.clone();
            broken["named_ranges"][0]["target"] = serde_json::to_value(target).unwrap();
            assert!(import_any(&broken.to_string())
                .unwrap_err()
                .contains("Invalid named range target"));
        }
        let mut excessive = json.clone();
        excessive["named_ranges"] =
            serde_json::to_value(vec![NamedRange::cell("Choice", 0, 0, 0); 8193]).unwrap();
        assert!(import_any(&excessive.to_string())
            .unwrap_err()
            .contains("definition limit"));
    }
}
