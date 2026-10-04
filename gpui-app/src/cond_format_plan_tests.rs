use super::*;
use crate::{history::History, table_edit::tests::fixture};
use visigrid_engine::cell::CellStyle;

fn range(r1: usize, c1: usize, r2: usize, c2: usize) -> CellRange {
    CellRange {
        start_row: r1,
        start_col: c1,
        end_row: r2,
        end_col: c2,
    }
}
fn style() -> CondStyle {
    CondStyle::Named(CellStyle::Warning)
}
fn selected(wb: &Workbook, ranges: &[Range]) -> (Vec<CellRange>, (usize, usize)) {
    let sheet = wb.active_sheet();
    let rows = sheet
        .build_saved_table_view(sheet.rows)
        .unwrap()
        .unwrap()
        .rows()
        .clone();
    targets(
        sheet,
        &rows,
        Some(&sheet.manual_hidden_rows()),
        None,
        ranges,
    )
    .unwrap()
}
fn snapshot(store: &CondFormatStore) -> Vec<CondFormatRule> {
    store.iter().cloned().collect()
}
fn action(before: &CondFormatStore, after: &CondFormatStore) -> UndoAction {
    let mut actions = vec![UndoAction::CondFormatsCleared {
        sheet_index: 0,
        rules: snapshot(before),
    }];
    actions.extend(
        after
            .iter()
            .cloned()
            .map(|rule| UndoAction::CondFormatAdded {
                sheet_index: 0,
                rule,
            }),
    );
    UndoAction::Group {
        actions,
        description: "Conditional formatting".into(),
    }
}

#[test]
fn filtered_creation_rebases_relative_references_from_first_displayed_record() {
    let wb = fixture(true);
    let (ranges, anchor) = selected(&wb, &[((3, 1), (6, 3))]);
    assert_eq!(anchor, (5, 1));
    assert_eq!(ranges, vec![range(3, 1, 3, 3), range(5, 1, 6, 3)]);
    let draft = Draft::new(&wb, 0, ranges, anchor, None);
    let store = draft.build("=($C6+$C$4)*2>100", style()).unwrap();
    for r in [3, 5, 6] {
        for c in 1..=3 {
            let rule = store.iter().find(|rule| rule.covers(r, c)).unwrap();
            assert_eq!(rule.matches(r, c, wb.active_sheet()), r != 5);
        }
    }
    assert!(!store.any_rule_covers(4, 2));
    for rule in store.iter() {
        validate_rule(wb.active_sheet(), rule).unwrap();
    }
    assert_eq!(store.iter().next().unwrap().predicate, "=($C4+$C$4)*2>100");
}

#[test]
fn drafts_leave_persistence_revision_and_live_rules_unchanged() {
    let mut wb = fixture(true);
    let id = wb
        .active_sheet_mut()
        .cond_formats
        .add(vec![range(3, 1, 6, 3)], "=$C4>25", style());
    let json = visigrid_io::json::export_workbook(&wb, &[], 0).unwrap();
    let rev = wb.revision();
    let original = wb.active_sheet().cond_formats.get(id).unwrap().clone();
    let mut draft = Draft::new(
        &wb,
        0,
        original.ranges.clone(),
        (3, 1),
        Some((0, original.clone())),
    );
    draft.preview = Some(draft.build("=$C4>35", style()).unwrap());
    assert_eq!(wb.active_sheet().cond_formats.get(id), Some(&original));
    assert_eq!(wb.revision(), rev);
    assert_eq!(
        visigrid_io::json::export_workbook(&wb, &[], 0).unwrap(),
        json
    );
    assert!(draft.validate(&wb).is_ok());
    drop(draft);
    assert_eq!(wb.active_sheet().cond_formats.get(id), Some(&original));
}

#[test]
fn stale_drafts_detect_revision_sheet_identity_and_untracked_rule_changes() {
    let wb = fixture(true);
    let draft = Draft::new(&wb, 0, vec![range(3, 1, 6, 3)], (3, 1), None);
    let mut changed = wb.clone();
    changed.bump_revision_for_structure();
    assert!(draft
        .validate(&changed)
        .unwrap_err()
        .contains("Copy your draft"));
    let mut changed = wb.clone();
    changed
        .active_sheet_mut()
        .cond_formats
        .add(vec![range(0, 0, 0, 0)], "=TRUE", style());
    assert!(draft.validate(&changed).is_err());
    let replaced = Workbook::from_sheets(vec![Sheet::new(SheetId(999), 30, 8)], 0);
    assert!(draft.validate(&replaced).is_err());
}

#[test]
fn edited_disabled_rule_keeps_id_priority_and_original_per_range_anchors() {
    let mut wb = fixture(true);
    let store = &mut wb.active_sheet_mut().cond_formats;
    store.add(vec![range(0, 0, 0, 0)], "=TRUE", style());
    let id = store.add(
        vec![range(3, 1, 3, 3), range(5, 1, 6, 3)],
        "=$C4>20",
        style(),
    );
    store.get_mut(id).unwrap().enabled = false;
    store.add(vec![range(1, 0, 1, 0)], "=FALSE", style());
    let rule = wb.active_sheet().cond_formats.get(id).unwrap().clone();
    let draft = Draft::new(&wb, 0, rule.ranges.clone(), (3, 1), Some((1, rule)));
    let after = draft.build("=$C4>30", style()).unwrap();
    let edited = after.iter().nth(1).unwrap();
    assert_eq!(edited.id, id);
    assert!(!edited.enabled);
    assert_eq!(edited.predicate_at(5, 1).as_deref(), Some("=$C4>30"));
    assert_eq!(after.len(), 3);
}

#[test]
fn manual_hides_columns_and_overlapping_selections_are_deduplicated() {
    let wb = fixture(true);
    let sheet = wb.active_sheet();
    let rows = sheet
        .build_saved_table_view(sheet.rows)
        .unwrap()
        .unwrap()
        .rows()
        .clone();
    let (ranges, anchor) = targets(
        sheet,
        &rows,
        Some(&[3].into()),
        Some(&[2].into()),
        &[((3, 1), (6, 3)), ((3, 1), (6, 3))],
    )
    .unwrap();
    assert_eq!(anchor, (5, 1));
    assert_eq!(ranges, vec![range(5, 1, 6, 1), range(5, 3, 6, 3)]);
}

#[test]
fn clear_preserves_hidden_records_precedence_and_relative_anchors() {
    let wb = fixture(true);
    let mut before = CondFormatStore::new();
    before.add(vec![range(0, 0, 0, 0)], "=TRUE", style());
    let id = before.add(vec![range(3, 1, 6, 3)], "=($C4+$C$4)*2>0", style());
    before.add(
        vec![range(3, 1, 6, 3)],
        "=$C4>25",
        CondStyle::Named(CellStyle::Success),
    );
    let (cuts, _) = selected(&wb, &[((3, 1), (6, 3))]);
    let after = clear(&before, &cuts).unwrap();
    assert!(after.get(id).unwrap().covers(4, 2));
    assert_eq!(after.get(id).unwrap().predicate, "=($C5+$C$4)*2>0");
    assert_eq!(snapshot(&before)[0], snapshot(&after)[0]);
    for r in 0..8 {
        for c in 0..5 {
            let removed = cuts.iter().any(|cut| cut.contains(r, c));
            let actual = after.override_for_cell(r, c, wb.active_sheet());
            if removed {
                assert!(actual.is_none());
            } else {
                assert_eq!(actual, before.override_for_cell(r, c, wb.active_sheet()));
            }
        }
    }
}

#[test]
fn clearing_overlapping_ranges_preserves_first_range_anchor_and_disabled_state() {
    let mut wb = fixture(false);
    for row in 8..14 {
        for col in 0..6 {
            wb.set_cell_value_tracked(0, row, col, &(row * 10 + col).to_string());
        }
    }
    let mut before = CondFormatStore::new();
    let id = before.add(
        vec![range(8, 0, 12, 3), range(10, 1, 13, 5)],
        "=A9>105",
        style(),
    );
    let cuts = [range(9, 1, 11, 2)];
    let after = clear(&before, &cuts).unwrap();
    for row in 8..14 {
        for col in 0..6 {
            if cuts[0].contains(row, col) {
                assert!(!after.any_rule_covers(row, col));
            } else {
                assert_eq!(
                    after.override_for_cell(row, col, wb.active_sheet()),
                    before.override_for_cell(row, col, wb.active_sheet())
                );
            }
        }
    }
    before.get_mut(id).unwrap().enabled = false;
    assert!(clear(&before, &cuts).unwrap().iter().all(|r| !r.enabled));
}

#[test]
fn unsafe_adjacent_metadata_refuses_while_titles_headers_footers_and_other_sheets_work() {
    let wb = fixture(true);
    for r in [
        range(0, 0, 0, 7),
        range(2, 0, 2, 7),
        range(7, 0, 7, 7),
        range(3, 1, 6, 3),
    ] {
        validate_rule(
            wb.active_sheet(),
            &CondFormatRule::new(0, vec![r], "=TRUE", style()),
        )
        .unwrap();
    }
    let rule = CondFormatRule::new(0, vec![range(3, 0, 6, 1)], "=TRUE", style());
    assert!(validate_rule(wb.active_sheet(), &rule)
        .unwrap_err()
        .contains("beside"));
    validate_rule(&Sheet::new(SheetId(88), 30, 8), &rule).unwrap();
}

#[test]
fn grouped_history_preflights_all_rules_and_allows_canonical_hidden_restoration() {
    let wb = fixture(true);
    let safe = CondFormatRule::new(0, vec![range(4, 1, 4, 3)], "=$C5>0", style());
    let unsafe_rule = CondFormatRule::new(1, vec![range(4, 0, 4, 3)], "=TRUE", style());
    let group = UndoAction::Group {
        actions: vec![
            UndoAction::CondFormatAdded {
                sheet_index: 0,
                rule: safe.clone(),
            },
            UndoAction::CondFormatAdded {
                sheet_index: 0,
                rule: unsafe_rule.clone(),
            },
        ],
        description: "Group".into(),
    };
    assert!(validate_history(&wb, &group, true).is_err());
    validate_history(&wb, &group, false).unwrap();
    validate_history(
        &wb,
        &UndoAction::CondFormatsCleared {
            sheet_index: 0,
            rules: vec![safe],
        },
        false,
    )
    .unwrap();
    assert!(validate_history(
        &wb,
        &UndoAction::CondFormatsCleared {
            sheet_index: 0,
            rules: vec![unsafe_rule]
        },
        false
    )
    .is_err());
    assert!(wb.active_sheet().cond_formats.is_empty());
    assert!(crate::table_command_scope::metadata_history_allowed(
        &wb, &group
    ));
    let mut history = History::new();
    history.record_action_with_provenance(group, None);
    assert!(history
        .build_workbook_before(1, Some(&wb), 100, 10_000)
        .is_err());
}

#[test]
fn history_rewind_and_json_roundtrip_keep_hidden_rules_and_totals_values() {
    let mut base = fixture(true);
    let table = base.active_sheet().tables()[0].id;
    base.set_table_totals_visible(table, true, Default::default())
        .unwrap();
    let id = base
        .active_sheet_mut()
        .cond_formats
        .add(vec![range(3, 1, 7, 3)], "=$C4>0", style());
    let (cuts, _) = selected(&base, &[((3, 1), (6, 3))]);
    let after = clear(&base.active_sheet().cond_formats, &cuts).unwrap();
    let action = action(&base.active_sheet().cond_formats, &after);
    for direction in [true, false] {
        validate_history(&base, &action, direction).unwrap();
    }
    let mut history = History::new();
    history.record_action_with_provenance(action, None);
    for position in [0, 1] {
        let preview = history
            .build_workbook_before(position, Some(&base), 100, 10_000)
            .unwrap()
            .workbook;
        assert_eq!(
            snapshot(&preview.active_sheet().cond_formats),
            if position == 1 {
                snapshot(&after)
            } else {
                snapshot(&base.active_sheet().cond_formats)
            }
        );
        assert_eq!(
            preview.active_sheet().get_raw(7, 3),
            base.active_sheet().get_raw(7, 3)
        );
        assert_eq!(
            preview.active_sheet().table_view_spec(),
            base.active_sheet().table_view_spec()
        );
        let json = visigrid_io::json::export_workbook(&preview, &[], 0).unwrap();
        let loaded = visigrid_io::json::import_any(&json).unwrap().0;
        assert_eq!(
            snapshot(&loaded.active_sheet().cond_formats),
            snapshot(&preview.active_sheet().cond_formats)
        );
        assert!(loaded.active_sheet().cond_formats.get(id).is_some());
    }
}

#[test]
fn oversized_empty_and_out_of_bounds_selections_refuse_without_a_partial_plan() {
    let sheet = Sheet::new(SheetId(4), 2000, 100);
    let rows = RowView::new(2000);
    assert!(targets(&sheet, &rows, None, None, &[((0, 0), (1999, 99))])
        .unwrap_err()
        .contains("100,000"));
    assert!(targets(&sheet, &rows, Some(&[0].into()), None, &[((0, 0), (0, 0))]).is_err());
    assert!(targets(&sheet, &rows, None, None, &[((0, 0), (2000, 0))]).is_err());
}

#[test]
fn rebasing_does_not_silently_destroy_out_of_bounds_relative_references() {
    let wb = fixture(true);
    let (ranges, anchor) = selected(&wb, &[((3, 1), (6, 3))]);
    let draft = Draft::new(&wb, 0, ranges, anchor, None);
    assert!(draft
        .build("=A1>0", style())
        .unwrap_err()
        .contains("outside"));
    assert!(draft.build("=$A$1>0", style()).is_ok());
    assert!(wb.active_sheet().cond_formats.is_empty());
}

fn replay(wb: &mut Workbook, action: &UndoAction, forward: bool) {
    match action {
        UndoAction::CondFormatAdded { sheet_index, rule } => {
            apply(wb, *sheet_index, std::slice::from_ref(rule), forward)
        }
        UndoAction::CondFormatsCleared { sheet_index, rules } => {
            apply(wb, *sheet_index, rules, !forward)
        }
        UndoAction::Group { actions, .. } => {
            if forward {
                for action in actions {
                    replay(wb, action, true);
                }
            } else {
                for action in actions.iter().rev() {
                    replay(wb, action, false);
                }
            }
        }
        _ => unreachable!(),
    }
}

#[test]
fn undo_redo_restore_exact_rule_order_even_after_filter_changes() {
    let mut wb = fixture(true);
    let before = &mut wb.active_sheet_mut().cond_formats;
    before.add(vec![range(0, 0, 0, 0)], "=TRUE", style());
    before.add(vec![range(3, 1, 6, 3)], "=$C4>0", style());
    before.add(
        vec![range(3, 1, 6, 3)],
        "=$C4>25",
        CondStyle::Named(CellStyle::Success),
    );
    let before = wb.active_sheet().cond_formats.clone();
    let (cuts, _) = selected(&wb, &[((3, 1), (6, 3))]);
    let after = clear(&before, &cuts).unwrap();
    let action = action(&before, &after);
    for forward in [true, false, true, false] {
        validate_history(&wb, &action, forward).unwrap();
        let revision = wb.revision();
        replay(&mut wb, &action, forward);
        assert!(wb.revision() > revision);
        assert_eq!(
            snapshot(&wb.active_sheet().cond_formats),
            snapshot(if forward { &after } else { &before })
        );
        wb.active_sheet().build_saved_table_view(30).unwrap();
        wb.set_table_view_spec(wb.active_sheet_id(), None).unwrap();
    }
}

#[test]
fn fragmented_selection_and_clear_refuse_before_publishing_any_rules() {
    let sheet = Sheet::new(SheetId(1), 9000, 2);
    let rows = RowView::new(sheet.rows);
    let hidden: BTreeSet<_> = (0..9000).filter(|r| r % 2 == 1).collect();
    assert!(
        targets(&sheet, &rows, Some(&hidden), None, &[((0, 0), (8999, 0))])
            .unwrap_err()
            .contains("too many")
    );
    let mut store = CondFormatStore::new();
    store.add(vec![range(0, 0, 8999, 0)], "=A1>0", style());
    let original = snapshot(&store);
    let cuts: Vec<_> = hidden.iter().map(|&r| range(r, 0, r, 0)).collect();
    assert!(clear(&store, &cuts).unwrap_err().contains("too many"));
    assert_eq!(snapshot(&store), original);
}
