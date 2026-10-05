use super::*;
use crate::{
    history::{History, UndoAction},
    table_edit::tests::fixture,
};
use visigrid_engine::{
    named_range::NamedRangeTarget,
    sheet::{Sheet, SheetId},
};

fn rows(wb: &Workbook) -> RowView {
    wb.active_sheet()
        .build_saved_table_view(wb.active_sheet().rows)
        .unwrap()
        .map(|v| v.rows().clone())
        .unwrap_or_else(|| RowView::new(wb.active_sheet().rows))
}

#[test]
fn extraction_uses_sorted_record_and_skips_hidden_formulas() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 3, 3, "=C6+1");
    wb.set_cell_value_tracked(0, 4, 3, "=C6+2"); // Hidden East record.
    let view = rows(&wb);
    let draft = ExtractionDraft::capture(&wb, &view, (4, 3)).unwrap();
    assert_eq!(draft.literal, "C6"); // Slot 4 is canonical row 5, not hidden row 4.
    assert_eq!(draft.cells(), [(3, 3), (5, 3)]);
    assert_eq!(draft.occurrences, 2);
    let (candidate, commit) = draft
        .prepare(&wb, &view, "Amount", Some("Extracted input".into()))
        .unwrap();
    assert!(wb.get_named_range("Amount").is_none());
    assert_eq!(
        candidate.get_named_range("Amount").unwrap().target,
        NamedRangeTarget::Cell {
            sheet: 0,
            row: 5,
            col: 2
        }
    );
    assert_eq!(candidate.active_sheet().get_raw(3, 3), "=Amount+1");
    assert_eq!(candidate.active_sheet().get_raw(4, 3), "=C6+2");
    assert_eq!(candidate.active_sheet().get_raw(5, 3), "=Amount*2");
    assert_eq!(candidate.active_sheet().get_display(5, 3), "40");
    assert_eq!(
        candidate.active_sheet().table_view_spec(),
        wb.active_sheet().table_view_spec()
    );
    let base = wb.clone();
    wb.restore_snapshot_monotonic(&candidate);
    commit.replay(&mut wb, true).unwrap();
    assert!(wb.get_named_range("Amount").is_none());
    assert_eq!(wb.active_sheet().get_raw(3, 3), "=C6+1");
    commit.replay(&mut wb, false).unwrap();
    assert_eq!(wb.active_sheet().get_raw(5, 3), "=Amount*2");
    let mut history = History::new();
    history.record_named_range_action(UndoAction::TableBatchChanged {
        sheet_index: 0,
        commit: Box::new(commit),
        description: "Extract Amount".into(),
    });
    let rewind = history
        .build_workbook_before(1, Some(&base), 100, 10_000)
        .unwrap()
        .workbook;
    assert_eq!(rewind.active_sheet().get_raw(5, 3), "=Amount*2");
    assert_eq!(
        rewind.get_named_range("Amount"),
        candidate.get_named_range("Amount")
    );
}

#[test]
fn other_sheet_criteria_and_quoted_targets_keep_sheet_identity_and_roundtrip() {
    let mut wb = fixture(true);
    let other = wb.add_sheet_named("Other! O'Brien").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "3");
    wb.set_cell_value_tracked(other, 1, 0, "7");
    wb.set_cell_value_tracked(other, 3, 1, "=SUM('Other! O''Brien'!$A$1:$A$2)");
    wb.set_active_sheet(other);
    let view = RowView::new(wb.active_sheet().rows);
    let draft = ExtractionDraft::capture(&wb, &view, (3, 1)).unwrap();
    let (mut candidate, commit) = draft.prepare(&wb, &view, "Inputs", None).unwrap();
    assert_eq!(candidate.sheet(other).unwrap().get_display(3, 1), "10");
    assert_eq!(
        candidate.sheet(other).unwrap().get_raw(3, 1),
        "=SUM(Inputs)"
    );
    assert_eq!(
        candidate.get_named_range("Inputs").unwrap().target,
        NamedRangeTarget::Range {
            sheet: other,
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 0
        }
    );
    commit.replay(&mut candidate, true).unwrap();
    assert!(candidate.get_named_range("Inputs").is_none());
    commit.replay(&mut candidate, false).unwrap();
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("extracted.sheet");
    visigrid_io::native::save_workbook(&candidate, &path).unwrap();
    let loaded = visigrid_io::native::load_workbook(&path).unwrap();
    assert_eq!(
        loaded.get_named_range("Inputs"),
        candidate.get_named_range("Inputs")
    );
    assert_eq!(loaded.sheet(other).unwrap().get_display(3, 1), "10");
}

#[test]
fn changed_dialog_hidden_targets_and_collisions_never_publish() {
    let wb = fixture(true);
    let view = rows(&wb);
    let draft = ExtractionDraft::capture(&wb, &view, (4, 3)).unwrap();
    assert!(draft.prepare(&wb, &view, "Sales", None).is_err());
    let mut changed = wb.clone();
    changed.set_cell_value_tracked(0, 15, 1, "new");
    assert!(draft.prepare(&changed, &view, "Amount", None).is_err());
    let mut changed = wb.clone();
    changed.active_sheet_mut().set_value(5, 3, "=C6+3");
    assert!(draft.prepare(&changed, &view, "Amount", None).is_err());
    let mut changed = wb.clone();
    changed.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(draft.prepare(&changed, &view, "Amount", None).is_err());
    assert!(ExtractionDraft::capture(&changed, &view, (4, 3)).is_err());
    assert!(ExtractionDraft::capture(&wb, &view, (3, 3)).is_err());
    assert!(wb.get_named_range("Amount").is_none());
}

#[test]
fn formula_errors_and_unsafe_new_spills_refuse_the_whole_extraction() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 15, 1, "=SUM(C4:C7)");
    wb.set_cell_value_tracked(0, 16, 1, "=LET(Inputs,7,SUM(C4:C7)*Inputs)");
    let view = rows(&wb);
    let draft = ExtractionDraft::capture(&wb, &view, (15, 1)).unwrap();
    assert!(draft
        .prepare(&wb, &view, "Inputs", None)
        .unwrap_err()
        .contains("LET or LAMBDA"));
    assert!(wb.get_named_range("Inputs").is_none());
    wb.set_cell_value_tracked(0, 16, 1, "=SUM(C4:C7)");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(ROWS(Inputs)+2)"); // Spill beside the projected Table.
    let draft = ExtractionDraft::capture(&wb, &view, (15, 1)).unwrap();
    assert!(draft.prepare(&wb, &view, "Inputs", None).is_err());
    assert!(wb.get_named_range("Inputs").is_none());
    assert_eq!(wb.active_sheet().get_raw(15, 1), "=SUM(C4:C7)");
}

#[test]
fn manual_calculation_and_spill_formulas_recompute_as_one_extraction() {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(1), 30, 8)], 0);
    wb.set_cell_value_tracked(0, 0, 0, "3");
    wb.set_cell_value_tracked(0, 1, 0, "7");
    wb.set_cell_value_tracked(0, 0, 2, "=SORT(A1:A2)");
    assert_eq!(wb.active_sheet().get_display(1, 2), "7");
    wb.set_auto_recalc(false);
    let view = rows(&wb);
    let draft = ExtractionDraft::capture(&wb, &view, (0, 2)).unwrap();
    let (candidate, _) = draft.prepare(&wb, &view, "Inputs", None).unwrap();
    assert!(!candidate.auto_recalc());
    assert_eq!(candidate.active_sheet().get_raw(0, 2), "=SORT(Inputs)");
    assert_eq!(candidate.active_sheet().get_display(1, 2), "7");
    assert_eq!(
        candidate.get_named_range("Inputs").unwrap().target,
        NamedRangeTarget::Range {
            sheet: 0,
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 0
        }
    );
}

#[test]
fn literal_formula_text_and_protected_totals_are_not_ordinary_formula_targets() {
    let mut wb = fixture(true);
    wb.set_cell_text_tracked(0, 15, 1, "=SUM(C4:C7)");
    assert!(ExtractionDraft::capture(&wb, &rows(&wb), (15, 1)).is_err());
    let id = wb.active_sheet().tables()[0].id;
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_table_total(
        id,
        3,
        visigrid_engine::table::TableTotal {
            function: Some("custom".into()),
            formula: Some("=SUM(C4:C7)".into()),
            label: None,
        },
    )
    .unwrap();
    assert!(ExtractionDraft::capture(&wb, &rows(&wb), (0, 1)).is_err());
    assert_eq!(wb.active_sheet().get_raw(0, 1), "=SUM(C4:C7)");
    assert!(wb.list_named_ranges().is_empty());
}
