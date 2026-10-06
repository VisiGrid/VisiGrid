use visigrid_engine::{
    named_range::NamedRangeTarget,
    sheet::{MergedRegion, NUM_COLS, NUM_ROWS},
    structural::Axis,
    validation::{CellRange, ValidationRule},
    workbook::Workbook,
};
use visigrid_io::{json, native, xlsx};

#[test]
fn deleted_sheet_names_remain_ref_errors_through_all_save_paths() {
    let mut wb = Workbook::new();
    let input = wb.add_sheet_named("Inputs").unwrap();
    wb.define_name_for_cell("TaxPct", input, 0, 0).unwrap();
    wb.set_cell_value_tracked(input, 0, 0, "0.2");
    wb.set_cell_value_tracked(0, 0, 0, "=TaxPct*100");
    let (deleted, _) = wb.prepare_sheet_delete(wb.sheets()[input].id).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("names.sheet");
    let check = |wb: &Workbook| {
        assert_eq!(
            wb.get_named_range("TaxPct").unwrap().target,
            NamedRangeTarget::RefError
        );
        assert_eq!(wb.active_sheet().get_display(0, 0), "#REF!");
        assert!(wb.read_only_reason().is_none());
    };
    for writer in 0..3 {
        match writer {
            0 => native::save_workbook(&deleted, &path).unwrap(),
            1 => native::save_workbook_with_metadata(&deleted, &Default::default(), &path).unwrap(),
            _ => {
                native::save_workbook_full(&deleted, &Default::default(), &[], &[], &path).unwrap()
            }
        }
        check(&native::load_workbook(&path).unwrap());
    }
    check(
        &json::import_any(&json::export_workbook(&deleted, &[], 0).unwrap())
            .unwrap()
            .0,
    );
    let path = dir.path().join("names.xlsx");
    xlsx::export(&deleted, &path, None).unwrap();
    check(&xlsx::import(&path).unwrap().0);
}

#[test]
fn row_and_column_near_edge_metadata_origins_and_exclusions_survive_reopen() {
    for axis in [Axis::Row, Axis::Col] {
        let mut wb = Workbook::new();
        let (merge, range, origin, excluded, shifted_origin) = if axis == Axis::Row {
            (
                MergedRegion::new(NUM_ROWS - 3, 0, NUM_ROWS - 2, 1),
                CellRange::new(0, 2, NUM_ROWS - 2, 2),
                (0, 2),
                CellRange::single(NUM_ROWS - 3, 2),
                (1, 2),
            )
        } else {
            (
                MergedRegion::new(0, NUM_COLS - 3, 1, NUM_COLS - 2),
                CellRange::new(2, 0, 2, NUM_COLS - 2),
                (2, 0),
                CellRange::single(2, NUM_COLS - 3),
                (2, 1),
            )
        };
        wb.set_cell_value_tracked(0, merge.start.0, merge.start.1, "Edge title");
        wb.active_sheet_mut().add_merge(merge).unwrap();
        let mut rule = ValidationRule::custom("=A1>0");
        rule.reference_origin = Some(origin);
        wb.active_sheet_mut().validations.set(range, rule);
        wb.active_sheet_mut().validations.exclude(excluded);
        wb.structural_edit(0, axis, 0, 1, false).unwrap();
        let expected_merges = wb.active_sheet().merged_regions.clone();
        assert_eq!(expected_merges.len(), 1);
        assert!(expected_merges
            .iter()
            .all(|m| m.end.0 < NUM_ROWS && m.end.1 < NUM_COLS));
        assert_eq!(
            wb.active_sheet()
                .validations
                .iter()
                .next()
                .unwrap()
                .1
                .reference_origin,
            Some(shifted_origin)
        );
        let shifted_exclusion = if axis == Axis::Row {
            CellRange::single(NUM_ROWS - 2, 2)
        } else {
            CellRange::single(2, NUM_COLS - 2)
        };
        assert_eq!(
            wb.active_sheet()
                .validations
                .exclusions_iter()
                .copied()
                .collect::<Vec<_>>(),
            [shifted_exclusion]
        );
        let visigrid_engine::validation::ValidationType::Custom(source) = &wb
            .active_sheet()
            .validations
            .iter()
            .next()
            .unwrap()
            .1
            .rule_type
        else {
            panic!()
        };
        assert_eq!(source, if axis == Axis::Row { "=A2>0" } else { "=B1>0" });
        let expected_rules = serde_json::json!([
            wb.active_sheet().validations.iter().collect::<Vec<_>>(),
            wb.active_sheet()
                .validations
                .exclusions_iter()
                .collect::<Vec<_>>()
        ]);
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("edge.sheet");
        native::save_workbook(&wb, &path).unwrap();
        for loaded in [
            native::load_workbook(&path).unwrap(),
            json::import_any(&json::export_workbook(&wb, &[], 0).unwrap())
                .unwrap()
                .0,
        ] {
            assert!(
                loaded.read_only_reason().is_none(),
                "{:?}",
                loaded.read_only_reason()
            );
            assert_eq!(loaded.active_sheet().merged_regions, expected_merges);
            let (row, col) = expected_merges[0].start;
            assert_eq!(loaded.active_sheet().get_raw(row, col), "Edge title");
            assert_eq!(
                serde_json::json!([
                    loaded.active_sheet().validations.iter().collect::<Vec<_>>(),
                    loaded
                        .active_sheet()
                        .validations
                        .exclusions_iter()
                        .collect::<Vec<_>>()
                ]),
                expected_rules
            );
        }
        xlsx::export_to_buffer(&wb, None).unwrap();
    }
}
