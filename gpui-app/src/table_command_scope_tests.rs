use super::{metadata_history_allowed, restore_freeze_panes, sheet_metadata_allowed};
use crate::{
    history::{CellFormatPatch, CommentPatch, FormatActionKind, History, UndoAction},
    table_edit::tests::fixture,
    table_structure::{StructureLayout, TableStructureHistory},
};
use visigrid_engine::{
    cell::{CellComment, CellFormat, CellStyle},
    cond_format::{CondFormatRule, CondStyle},
    sheet::{Sheet, SheetId},
    structural::Axis,
    validation::CellRange,
    workbook::{StructureStep, Workbook},
};

fn workbook() -> Workbook {
    let mut wb = fixture(true);
    assert!(wb.restore_sheet(1, Sheet::new_with_name(SheetId(99), 30, 8, "Report")));
    wb
}

fn bold(sheet_index: usize) -> UndoAction {
    UndoAction::Format {
        sheet_index,
        patches: vec![CellFormatPatch { remove_cell_on_undo: false,
            row: 0,
            col: 0,
            before: CellFormat::default(),
            after: CellFormat {
                bold: true,
                ..Default::default()
            },
        }],
        kind: FormatActionKind::Bold,
        description: "Bold title".into(),
    }
}

fn freeze() -> UndoAction {
    UndoAction::FreezePanesChanged {
        sheet_id: SheetId(99),
        old_frozen_rows: 0,
        old_frozen_cols: 0,
        new_frozen_rows: 1,
        new_frozen_cols: 1,
    }
}

#[test]
fn metadata_scope_follows_the_target_sheet_even_after_switching_tabs() {
    let mut wb = workbook();
    for active in [0, 1] {
        wb.set_active_sheet(active);
        assert!(sheet_metadata_allowed(&wb, 1));
        assert!(!sheet_metadata_allowed(&wb, 0));
        assert!(metadata_history_allowed(&wb, &bold(1)));
        assert!(metadata_history_allowed(&wb, &freeze()));
        assert!(metadata_history_allowed(&wb, &bold(0)));
    }
    assert!(!sheet_metadata_allowed(&wb, 999));
    assert!(!metadata_history_allowed(&wb, &bold(999)));
    let mut spec = wb.sheet(0).unwrap().table_view_spec().unwrap().clone();
    spec.clear_sort();
    wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
        .unwrap();
    assert!(
        !sheet_metadata_allowed(&wb, 0),
        "filter alone still protects the sheet"
    );
    spec.clear_filters();
    spec.show_filter_buttons = false;
    wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
    assert!(
        sheet_metadata_allowed(&wb, 0),
        "hidden arrows alone do not block metadata"
    );
}

#[test]
fn mixed_history_groups_cannot_bypass_the_criteria_guard() {
    let wb = workbook();
    let group = |actions| UndoAction::Group {
        actions,
        description: "Group".into(),
    };
    assert!(metadata_history_allowed(
        &wb,
        &group(vec![bold(1), freeze()])
    ));
    assert!(metadata_history_allowed(
        &wb,
        &group(vec![bold(1), bold(0)])
    ));
    assert!(!metadata_history_allowed(
        &wb,
        &group(vec![
            bold(1),
            UndoAction::Values {
                sheet_index: 1,
                changes: vec![],
            }
        ])
    ));
    assert!(!metadata_history_allowed(&wb, &group(vec![])));
}

#[test]
fn metadata_rewind_preserves_filtered_records_and_restores_the_other_sheet() {
    let base = workbook();
    let mut history = History::new();
    let action = UndoAction::Group {
        actions: vec![
            bold(1),
            UndoAction::Comments {
                sheet_index: 1,
                patches: vec![CommentPatch { remove_cell_on_undo: false,
                    row: 0,
                    col: 0,
                    before: None,
                    after: Some(CellComment {
                        text: "Report note".into(),
                        author: "Tester".into(),
                    }),
                }],
                description: "Comment".into(),
            },
            UndoAction::CondFormatAdded {
                sheet_index: 1,
                rule: CondFormatRule::new(
                    1,
                    vec![CellRange {
                        start_row: 0,
                        end_row: 0,
                        start_col: 0,
                        end_col: 0,
                    }],
                    "=TRUE",
                    CondStyle::Named(CellStyle::Note),
                ),
            },
            freeze(),
        ],
        description: "Style report".into(),
    };
    assert!(metadata_history_allowed(&base, &action));
    history.record_action_with_provenance(action, None);
    for position in [0, 1] {
        let result = history
            .build_workbook_before(position, Some(&base), 100, 10_000)
            .unwrap();
        let report = result.workbook.sheet(1).unwrap();
        assert_eq!(report.get_format(0, 0).bold, position == 1);
        assert_eq!(report.comment(0, 0).is_some(), position == 1);
        assert_eq!(report.cond_formats.iter().count(), position);
        assert_eq!(report.frozen_panes, (position, position));
        let table = result.workbook.sheet(0).unwrap();
        assert_eq!(
            table.table_view_spec(),
            base.sheet(0).unwrap().table_view_spec()
        );
        assert_eq!(table.get_raw(4, 1), "East");
        assert_eq!(table.frozen_panes, (0, 0));
        assert!(!table
            .build_saved_table_view(30)
            .unwrap()
            .unwrap()
            .rows()
            .is_data_row_visible(4));
    }
}

#[test]
fn freeze_undo_redo_changes_the_recorded_sheet_not_the_active_sheet() {
    let mut wb = workbook();
    wb.set_active_sheet(0);
    for frozen in [(1, 1), (0, 0), (1, 1)] {
        restore_freeze_panes(&mut wb, SheetId(99), frozen).unwrap();
        assert_eq!(wb.sheet(1).unwrap().frozen_panes, frozen);
        assert_eq!(wb.active_sheet().frozen_panes, (0, 0));
    }
    assert!(restore_freeze_panes(&mut wb, SheetId(1000), (2, 2)).is_err());
    assert_eq!(wb.active_sheet().frozen_panes, (0, 0));
}

#[test]
fn freeze_and_structural_history_rewind_together_on_an_unfiltered_sheet() {
    let base = workbook();
    let mut wb = base.clone();
    restore_freeze_panes(&mut wb, SheetId(99), (1, 1)).unwrap();
    let steps = vec![StructureStep {
        axis: Axis::Row,
        at: 0,
        count: 1,
        delete: false,
    }];
    let layout = StructureLayout::default();
    let after = layout.shifted(wb.sheet(1).unwrap(), &steps).unwrap();
    let (_, commit) = wb.prepare_guarded_structure(1, steps).unwrap();
    let mut history = History::new();
    history.record_action_with_provenance(freeze(), None);
    history.record_action_with_provenance(
        UndoAction::TableStructureChanged {
            sheet_index: 1,
            history: Box::new(TableStructureHistory {
                source_frozen: Some((1, 1)),
                commit,
                before: layout,
                after,
            }),
            description: "Insert row".into(),
        },
        None,
    );
    history.record_action_with_provenance(
        UndoAction::FreezePanesChanged {
            sheet_id: SheetId(99),
            old_frozen_rows: 2,
            old_frozen_cols: 1,
            new_frozen_rows: 0,
            new_frozen_cols: 0,
        },
        None,
    );
    for (position, frozen) in [(0, (0, 0)), (1, (1, 1)), (2, (2, 1)), (3, (0, 0))] {
        let result = history
            .build_workbook_before(position, Some(&base), 100, 10_000)
            .unwrap();
        assert_eq!(
            result.workbook.sheet(1).unwrap().frozen_panes,
            frozen,
            "position {position}"
        );
        assert_eq!(result.workbook.sheet(0).unwrap().frozen_panes, (0, 0));
    }
}

#[test]
fn a_missing_freeze_target_refuses_the_whole_history_group() {
    let wb = workbook();
    let action = UndoAction::Group {
        actions: vec![
            bold(1),
            UndoAction::FreezePanesChanged {
                sheet_id: SheetId(1000),
                old_frozen_rows: 0,
                old_frozen_cols: 0,
                new_frozen_rows: 1,
                new_frozen_cols: 0,
            },
        ],
        description: "Stale report".into(),
    };
    assert!(super::validate_freeze_history(&wb, &action).is_err());
    assert!(!metadata_history_allowed(&wb, &action));
    assert!(super::validate_freeze_history(&wb, &freeze()).is_ok());
}

#[test]
fn percent_format_preflight_detects_value_conversions_across_the_whole_selection() {
    use visigrid_engine::cell::NumberFormat;
    let mut wb = workbook();
    wb.set_cell_value_tracked(1, 0, 0, "0.25");
    wb.set_cell_text_tracked(1, 2, 0, "25%");
    let percent = NumberFormat::Percent { decimals: 2 };
    assert!(!super::number_format_changes_values(
        wb.sheet(1).unwrap(),
        &percent,
        &[((0, 0), (0, 0))]
    ));
    assert!(super::number_format_changes_values(
        wb.sheet(1).unwrap(),
        &percent,
        &[((0, 0), (0, 0)), ((2, 0), (2, 0))]
    ));
    assert!(!super::number_format_changes_values(
        wb.sheet(1).unwrap(),
        &NumberFormat::General,
        &[((0, 0), (2, 0))]
    ));
    assert_eq!(wb.sheet(1).unwrap().get_raw(2, 0), "25%");
    assert_eq!(super::percent_format_value("25%"), Some("0.25".into()));
    assert_eq!(super::percent_format_value("not a percent"), None);
}

#[test]
fn styling_a_cleared_view_does_not_block_edits_in_another_filtered_table() {
    let mut wb = workbook();
    for (row, text) in [(2, "Amount"), (3, "10"), (4, "20")] {
        wb.set_cell_value_tracked(1, row, 1, text);
    }
    let id = wb
        .create_table(
            SheetId(99),
            visigrid_engine::table::TableRange {
                start_row: 2,
                end_row: 4,
                start_col: 1,
                end_col: 1,
            },
            "Other",
        )
        .unwrap()
        .table_id();
    let mut spec = visigrid_engine::table_view::TableViewSpec::new(id);
    spec.show_filter_buttons = false;
    wb.set_table_view_spec(SheetId(99), Some(spec)).unwrap();
    let mut history = History::new();
    let mut formatting = bold(1);
    if let UndoAction::Format { patches, .. } = &mut formatting {
        patches[0].row = 3;
        patches[0].col = 6; // Beside the cleared Table's body.
    }
    history.record_action_with_provenance(formatting, None);
    let styled = history
        .build_workbook_before(1, Some(&wb), 100, 10_000)
        .unwrap()
        .workbook;
    assert!(styled
        .sheet(1)
        .unwrap()
        .build_saved_table_view(30)
        .unwrap()
        .is_none());
    let edited = crate::table_edit::prepare_table_writes(
        &styled,
        0,
        &[crate::table_edit::TableCellWrite::value(3, 2, "35".into())],
    )
    .unwrap();
    assert_eq!(edited.sheet(0).unwrap().get_raw(3, 2), "35");
    assert_eq!(edited.sheet(0).unwrap().get_raw(4, 1), "East");
    assert!(edited.sheet(1).unwrap().get_format(3, 6).bold);
}
