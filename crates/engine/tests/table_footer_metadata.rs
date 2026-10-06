use visigrid_engine::{
    cell::CellStyle,
    cond_format::{CondFormatRule, CondStyle},
    table::{TableId, TableRange},
    validation::{CellRange, ValidationRule, ValidationType},
    workbook::Workbook,
};

fn book() -> (Workbook, TableId) {
    let mut wb = Workbook::new();
    for (r, values) in [
        ["Group", "Amount"],
        ["West", "10"],
        ["East", "20"],
        ["West", "30"],
    ]
    .iter()
    .enumerate()
    {
        for (c, value) in values.iter().enumerate() {
            wb.set_cell_value_tracked(0, r, c, value);
        }
    }
    let id = wb
        .create_table(
            wb.active_sheet_id(),
            TableRange {
                start_row: 0,
                start_col: 0,
                end_row: 3,
                end_col: 1,
            },
            "Sales",
        )
        .unwrap()
        .table_id();
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    (wb, id)
}
fn custom(source: &str, origin: Option<(usize, usize)>) -> ValidationRule {
    let mut rule = ValidationRule::custom(source);
    rule.reference_origin = origin;
    rule
}
fn predicate(wb: &Workbook, row: usize, col: usize) -> Vec<String> {
    wb.active_sheet()
        .cond_formats
        .iter()
        .filter_map(|r| r.predicate_at(row, col))
        .collect()
}
fn constraint(wb: &Workbook, row: usize, col: usize) -> Option<String> {
    wb.active_sheet().validations.get(row, col).map(|r| {
        let ValidationType::Custom(source) = &r.at(row, col).rule_type else {
            panic!()
        };
        source.clone()
    })
}
fn metadata(wb: &Workbook) -> serde_json::Value {
    serde_json::json!([
        wb.active_sheet().cond_formats,
        wb.active_sheet().validations.iter().collect::<Vec<_>>(),
        wb.active_sheet()
            .validations
            .exclusions_iter()
            .collect::<Vec<_>>()
    ])
}

#[test]
fn footer_rules_replace_destination_coverage_and_restore_exactly() {
    let (mut wb, id) = book();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::single(4, 1)],
        "=$B$5>50",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(5, 0, 5, 2)],
        "=FALSE",
        CondStyle::Named(CellStyle::Error),
    );
    wb.active_sheet_mut()
        .validations
        .set(CellRange::single(4, 1), custom("=B5>0", Some((4, 1))));
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(5, 0, 5, 2), custom("=FALSE", None));
    let original = metadata(&wb);
    let commit = wb.append_table_rows(id, 1, &[(4, 1, "10".into())]).unwrap();
    assert!(predicate(&wb, 4, 1).is_empty());
    assert_eq!(predicate(&wb, 5, 1), ["=$B$6>50"]);
    assert_eq!(predicate(&wb, 5, 2), ["=FALSE"]);
    assert_eq!(constraint(&wb, 4, 1), None);
    assert_eq!(constraint(&wb, 5, 1).as_deref(), Some("=B6>0"));
    assert_eq!(
        constraint(&wb, 5, 0),
        None,
        "empty source metadata also replaces destination"
    );
    assert_eq!(constraint(&wb, 5, 2).as_deref(), Some("=FALSE"));
    let moved = metadata(&wb);
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(metadata(&wb), original);
    wb.apply_table_commit(&commit, false).unwrap();
    assert_eq!(metadata(&wb), moved);
}

#[test]
fn spanning_rules_keep_continuous_coverage_and_surrounding_anchors() {
    let (mut wb, id) = book();
    let ranges = vec![CellRange::new(0, 1, 8, 1), CellRange::new(2, 1, 10, 1)];
    let rule_id = wb.active_sheet_mut().cond_formats.add(
        ranges,
        "=H1+$B$5>0",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut()
        .cond_formats
        .get_mut(rule_id)
        .unwrap()
        .enabled = false;
    wb.active_sheet_mut().validations.set(
        CellRange::new(0, 1, 10, 1),
        custom("=H1+$B$5>0", Some((0, 1))),
    );
    let original = metadata(&wb);
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    for row in 0..=10 {
        let cf_old_row = if row <= 8 { row } else { row - 2 };
        let expected_cf = format!("=H{}+$B$6>0", cf_old_row + 1);
        assert_eq!(predicate(&wb, row, 1), [expected_cf], "CF row {row}");
        let expected = format!("=H{}+$B$6>0", row + 1);
        assert_eq!(constraint(&wb, row, 1), Some(expected), "validation row {row}");
    }
    assert!(wb.active_sheet().cond_formats.iter().all(|r| !r.enabled));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(metadata(&wb), original);
}

#[test]
fn relative_readers_elsewhere_only_follow_at_the_actual_footer_reference() {
    let (mut wb, id) = book();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(0, 4, 8, 4)],
        "=$B1>0",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(0, 5, 8, 5), custom("=$B1>0", Some((0, 5))));
    let other = wb.add_sheet_named("Other").unwrap();
    wb.sheet_mut(other).unwrap().cond_formats.add(
        vec![CellRange::new(0, 0, 8, 0)],
        "=Sheet1!$B1>0",
        CondStyle::Named(CellStyle::Warning),
    );
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    for row in 0..=8 {
        let referenced = if row == 4 { 6 } else { row + 1 };
        assert_eq!(predicate(&wb, row, 4), [format!("=$B{referenced}>0")]);
        assert_eq!(constraint(&wb, row, 5), Some(format!("=$B{referenced}>0")));
        let refs: Vec<_> = wb
            .sheet(other)
            .unwrap()
            .cond_formats
            .iter()
            .filter_map(|r| r.predicate_at(row, 0))
            .collect();
        assert_eq!(refs, [format!("=Sheet1!$B{referenced}>0")]);
    }
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(wb.active_sheet().cond_formats.len(), 1);
    assert_eq!(wb.active_sheet().validations.iter().count(), 1);
}

#[test]
fn validation_overlap_precedence_and_exclusions_follow_the_moved_cells() {
    let (mut wb, id) = book();
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(0, 0, 8, 0), custom("=TRUE", None));
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(1, 0, 10, 1), custom("=FALSE", None));
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::single(4, 1));
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::single(5, 1));
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::new(1, 5, 2, 6));
    let original = metadata(&wb);
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    assert!(!wb.active_sheet().validations.is_excluded(4, 1));
    assert!(wb.active_sheet().validations.is_excluded(5, 1));
    assert!(wb.active_sheet().validations.is_excluded(1, 5));
    for row in 0..=10 {
        for col in 0..=1 {
            let expected = if row == 5 && col == 1 || row == 0 && col == 1 {
                None
            } else if col == 0 && row <= 8 {
                Some("=TRUE".to_string())
            } else {
                Some("=FALSE".to_string())
            };
            assert_eq!(constraint(&wb, row, col), expected, "{row},{col}");
        }
    }
    wb.active_sheet_mut()
        .validations
        .remove_exclusion(&CellRange::single(5, 1));
    assert_eq!(
        constraint(&wb, 5, 1).as_deref(),
        Some("=FALSE"),
        "excluded definition was retained"
    );
    assert!(
        wb.apply_table_commit(&commit, true).is_err(),
        "changed rules make replay stale"
    );
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::single(5, 1));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(metadata(&wb), original);
}

#[test]
fn shrink_and_compound_width_moves_keep_released_column_rules() {
    let (mut wb, id) = book();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(4, 0, 4, 1)],
        "=B5>0",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(4, 0, 4, 1), custom("=B5>0", Some((4, 0))));
    wb.clear_cell_tracked(0, 3, 0);
    let original = metadata(&wb);
    let range = wb.table(id).unwrap().1.range;
    let commit = wb
        .resize_table(
            id,
            TableRange {
                end_row: 2,
                end_col: 0,
                ..range
            },
        )
        .unwrap();
    // Released B5 keeps its own rule and its references; surviving A5 moves to A4.
    assert_eq!(predicate(&wb, 3, 0), ["=B5>0"]);
    assert_eq!(predicate(&wb, 4, 1), ["=C5>0"]);
    assert_eq!(constraint(&wb, 3, 0).as_deref(), Some("=B5>0"));
    assert_eq!(constraint(&wb, 4, 1).as_deref(), Some("=C5>0"));
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(metadata(&wb), original);
}

#[test]
fn newly_introduced_rules_refuse_unlinked_history_and_bad_ranges_refuse_atomically() {
    let (mut wb, id) = book();
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    wb.active_sheet_mut()
        .validations
        .exclude(CellRange::single(5, 1));
    let rev = wb.revision();
    assert!(wb.apply_table_commit(&commit, true).is_err());
    assert_eq!(wb.revision(), rev);
    wb.active_sheet_mut().validations.clear_exclusions();
    wb.apply_table_commit(&commit, true).unwrap();
    wb.active_sheet_mut().cond_formats.insert_at(
        0,
        CondFormatRule::new(
            77,
            vec![CellRange::single(usize::MAX, 0)],
            "=TRUE",
            CondStyle::Named(CellStyle::Warning),
        ),
    );
    let rev = wb.revision();
    assert!(wb.append_table_rows(id, 1, &[]).is_err());
    assert_eq!(wb.revision(), rev);
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(4));
}

#[test]
fn excessive_fragmentation_refuses_without_publishing_rule_or_footer_changes() {
    let (mut wb, id) = book();
    for _ in 0..4097 {
        wb.active_sheet_mut().cond_formats.add(
            vec![CellRange::single(4, 1)],
            "=TRUE",
            CondStyle::Named(CellStyle::Warning),
        );
    }
    let rev = wb.revision();
    let error = wb.append_table_rows(id, 1, &[]).unwrap_err();
    assert!(error.contains("fragments"), "{error}");
    assert_eq!(wb.revision(), rev);
    assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(4));
    assert_eq!(wb.active_sheet().cond_formats.len(), 4097);
    assert!(wb
        .active_sheet()
        .cond_formats
        .iter()
        .all(|r| r.ranges == [CellRange::single(4, 1)]));
}

#[test]
fn untouched_definitions_and_relative_origin_before_grid_keep_their_meaning() {
    let (mut wb, id) = book();
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(0, 5, 9, 5), custom("=B5>0", Some((5, 5))));
    let untouched = custom("=TRUE", None);
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(0, 7, 9, 7), untouched.clone());
    wb.active_sheet_mut()
        .validations
        .set(CellRange::single(5, 7), custom("=FALSE", None));
    wb.append_table_rows(id, 1, &[]).unwrap();
    assert_eq!(constraint(&wb, 0, 5).as_deref(), Some("=#REF!>0"));
    for row in 1..=9 {
        let reference = if row == 5 { 6 } else { row };
        assert_eq!(constraint(&wb, row, 5), Some(format!("=B{reference}>0")));
    }
    let rules: Vec<_> = wb
        .active_sheet()
        .validations
        .iter()
        .filter(|(r, _)| r.start_col == 7)
        .collect();
    assert_eq!(
        rules.len(),
        2,
        "unrelated shadowed definition remains available"
    );
    assert_eq!(rules[0].1, &untouched);
}

#[test]
fn relative_rule_partitioning_preserves_both_row_and_column_offsets() {
    let (mut wb, id) = book();
    wb.active_sheet_mut().cond_formats.add(
        vec![CellRange::new(0, 5, 8, 8)],
        "=A1>0",
        CondStyle::Named(CellStyle::Warning),
    );
    wb.active_sheet_mut()
        .validations
        .set(CellRange::new(0, 5, 8, 8), custom("=A1>0", Some((0, 5))));
    wb.append_table_rows(id, 1, &[]).unwrap();
    for row in 0..=8 {
        for col in 5..=8 {
            let referenced_row = if row == 4 && col <= 6 { 6 } else { row + 1 };
            let referenced_col = char::from(b'A' + (col - 5) as u8);
            let expected = format!("={referenced_col}{referenced_row}>0");
            assert_eq!(predicate(&wb, row, col), [expected.clone()]);
            assert_eq!(constraint(&wb, row, col), Some(expected));
        }
    }
}

#[test]
fn newly_adjacent_footer_metadata_refuses_a_filtered_append_atomically() {
    use visigrid_engine::{
        filter::SortDirection,
        table_view::{TableSort, TableViewSpec},
    };
    for validation in [false, true] {
        let (mut wb, id) = book();
        let mut view = TableViewSpec::new(id);
        view.sort = Some(TableSort {
            column: wb.table(id).unwrap().1.columns[0].id,
            direction: SortDirection::Ascending,
        });
        wb.set_table_view_spec(wb.active_sheet_id(), Some(view))
            .unwrap();
        let range = CellRange::new(4, 0, 4, 2);
        if validation {
            wb.active_sheet_mut()
                .validations
                .set(range, custom("=TRUE", None));
        } else {
            wb.active_sheet_mut().cond_formats.add(
                vec![range],
                "=TRUE",
                CondStyle::Named(CellStyle::Warning),
            );
        }
        let before = metadata(&wb);
        let rev = wb.revision();
        assert!(wb.append_table_rows(id, 1, &[]).is_err());
        assert_eq!(wb.revision(), rev);
        assert_eq!(metadata(&wb), before);
        assert_eq!(wb.table(id).unwrap().1.totals_row(), Some(4));
    }
}

#[test]
fn fragment_ids_respect_persisted_rules_and_refuse_invalid_identity_stores() {
    let (mut wb, id) = book();
    let rule = CondFormatRule::new(
        100,
        vec![CellRange::new(0, 1, 8, 1)],
        "=B1>0",
        CondStyle::Named(CellStyle::Warning),
    );
    let store = |rules, next_id| {
        serde_json::from_value(serde_json::json!({"rules": rules, "next_id": next_id})).unwrap()
    };
    wb.active_sheet_mut().cond_formats = store(vec![rule.clone()], 0u64);
    wb.active_sheet_mut().cond_formats.reparse_all();
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    let ids: Vec<_> = wb
        .active_sheet()
        .cond_formats
        .iter()
        .map(|r| r.id)
        .collect();
    assert!(ids.len() > 1);
    assert_eq!(ids[0], 100);
    assert!(ids[1..].iter().all(|&id| id > 100));
    assert_eq!(
        ids.iter().collect::<std::collections::BTreeSet<_>>().len(),
        ids.len()
    );
    wb.apply_table_commit(&commit, true).unwrap();
    let mut exhausted = rule.clone();
    exhausted.id = u64::MAX;
    for rules in [vec![rule.clone(), rule.clone()], vec![exhausted]] {
        wb.active_sheet_mut().cond_formats = store(rules, 0u64);
        let rev = wb.revision();
        assert!(wb
            .append_table_rows(id, 1, &[])
            .unwrap_err()
            .contains("IDs"));
        assert_eq!(wb.revision(), rev);
    }
    wb.active_sheet_mut().cond_formats = store(vec![rule], u64::MAX);
    assert!(wb
        .append_table_rows(id, 1, &[])
        .unwrap_err()
        .contains("exhausted"));
}

#[test]
fn full_column_rules_remain_continuous_after_append_and_undo() {
    use visigrid_engine::sheet::NUM_ROWS;
    let (mut wb, id) = book();
    let column = CellRange::new(0, 1, NUM_ROWS - 1, 1);
    wb.active_sheet_mut().cond_formats.add(vec![column], "=B1>0", CondStyle::Named(CellStyle::Warning));
    wb.active_sheet_mut().validations.set(column, custom("=B1>0", Some((0, 1))));
    let original = metadata(&wb);
    let commit = wb.append_table_rows(id, 1, &[]).unwrap();
    for row in [0, 3, 4, 5, 6, NUM_ROWS - 1] {
        assert!(!predicate(&wb, row, 1).is_empty(), "CF row {row}");
        assert!(constraint(&wb, row, 1).is_some(), "validation row {row}");
    }
    wb.apply_table_commit(&commit, true).unwrap();
    assert_eq!(metadata(&wb), original);
}

#[test]
fn body_and_footer_rules_cover_two_appends_and_shrink() {
    for end in [4, 999] {
        let (mut wb, id) = book();
        // Leave the shrink destination empty; totals never overwrite data.
        wb.clear_cell_tracked(0, 3, 0);
        wb.clear_cell_tracked(0, 3, 1);
        // B3:B1000 is the common worksheet-wide dropdown; B3:B5 ends
        // exactly at this Table's footer and must grow with it as well.
        let range = CellRange::new(2, 1, end, 1);
        wb.active_sheet_mut().validations.set(range, ValidationRule::list_inline(vec!["Yes".into(), "No".into()]));
        wb.active_sheet_mut().cond_formats.add(vec![range], "=TRUE", CondStyle::Named(CellStyle::Warning));
        let original = metadata(&wb);
        let first = wb.append_table_rows(id, 1, &[]).unwrap();
        let second = wb.append_table_rows(id, 1, &[]).unwrap();
        for row in 2..=if end == 4 { 6 } else { 999 } {
            assert!(wb.active_sheet().validations.has_validation(row, 1), "missing dropdown at {row}");
            assert_eq!(predicate(&wb, row, 1), ["=TRUE"], "missing format at {row}");
        }
        let shrink = wb.resize_table(id, TableRange { start_row: 0, start_col: 0, end_row: 2, end_col: 1 }).unwrap();
        for row in 2..=if end == 4 { 3 } else { 999 } {
            assert!(wb.active_sheet().validations.has_validation(row, 1));
            assert_eq!(predicate(&wb, row, 1), ["=TRUE"]);
        }
        wb.apply_table_commit(&shrink, true).unwrap();
        wb.apply_table_commit(&second, true).unwrap();
        wb.apply_table_commit(&first, true).unwrap();
        assert_eq!(metadata(&wb), original);
    }
}
