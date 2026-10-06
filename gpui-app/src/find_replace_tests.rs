use super::plan::{self, MatchKind};
use crate::{
    history::{History, UndoAction},
    table_cell_history::TableCellsCommit,
    table_edit::{prepare_table_writes, tests::fixture},
};
use visigrid_engine::{
    cell::{CellComment, ValueRef},
    filter::RowView,
    sheet::{Sheet, SheetId},
    workbook::Workbook,
};

fn rows(wb: &Workbook, index: usize) -> RowView {
    let sheet = wb.sheet(index).unwrap();
    sheet
        .build_saved_table_view(sheet.rows)
        .unwrap()
        .map(|v| v.rows().clone())
        .unwrap_or_else(|| RowView::new(sheet.rows))
}

#[test]
fn search_follows_display_order_and_skips_filtered_and_manual_hidden_records() {
    let mut wb = fixture(true);
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "West").unwrap();
    assert_eq!(hits.iter().map(|h| h.row).collect::<Vec<_>>(), [5, 3, 6]);
    assert_eq!(
        hits.iter()
            .map(|h| plan::visible_row(&view, None, h.row).unwrap())
            .collect::<Vec<_>>(),
        [4, 5, 6]
    );
    assert!(plan::search(wb.active_sheet(), 0, &view, None, None, "East")
        .unwrap()
        .is_empty());
    wb.active_sheet_mut()
        .set_manual_hidden_rows([3].into())
        .unwrap();
    let hits = plan::search(wb.active_sheet(), 0, &rows(&wb, 0), None, None, "West").unwrap();
    assert_eq!(hits.iter().map(|h| h.row).collect::<Vec<_>>(), [5, 6]);
    wb.set_table_view_spec(wb.active_sheet_id(), None).unwrap();
    let hits = plan::search(
        wb.active_sheet(),
        0,
        &rows(&wb, 0),
        Some(&wb.active_sheet().manual_hidden_rows()),
        None,
        "West",
    )
    .unwrap();
    assert_eq!(hits.iter().map(|h| h.row).collect::<Vec<_>>(), [5, 6]);
}

#[test]
fn replace_all_can_hide_every_match_with_totals_and_sparse_undo_redo_rewind() {
    let mut base = fixture(true);
    let id = base.active_sheet().tables()[0].id;
    base.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    base.active_sheet_mut().set_comment(
        3,
        1,
        Some(CellComment {
            text: "Keep".into(),
            author: "QA".into(),
        }),
    );
    base.active_sheet_mut().set_bold(3, 1, true);
    let view = rows(&base, 0);
    let hits = plan::search(base.active_sheet(), 0, &view, None, None, "West").unwrap();
    let (writes, count) = plan::replacements(&base, 0, &view, None, None, &hits, "North").unwrap();
    assert_eq!(count, 3);
    let mut after = prepare_table_writes(&base, 0, &writes).unwrap();
    assert!(
        plan::search(after.active_sheet(), 0, &rows(&after, 0), None, None, "North")
            .unwrap()
            .is_empty()
    );
    assert_eq!(after.active_sheet().get_raw(4, 1), "East");
    assert_eq!(after.active_sheet().get_display(7, 3), "0");
    assert!(after.active_sheet().get_format(3, 1).bold);
    assert_eq!(after.active_sheet().comment(3, 1).unwrap().text, "Keep");
    let commit = TableCellsCommit::capture(
        base.active_sheet(),
        after.active_sheet(),
        writes.iter().map(|w| (w.row, w.col)),
    );
    assert_eq!(commit.patches.len(), 3);
    commit.replay(&mut after, true).unwrap();
    assert_eq!(after.active_sheet().get_display(7, 3), "180");
    commit.replay(&mut after, false).unwrap();
    assert_eq!(after.active_sheet().get_display(7, 3), "0");
    let mut history = History::new();
    history.record_action_with_provenance(
        UndoAction::TableCellsChanged {
            sheet_index: 0,
            commit: Box::new(commit),
            description: "Find and Replace".into(),
        },
        None,
    );
    let replay = history
        .build_workbook_before(1, Some(&base), 100, 10_000)
        .unwrap()
        .workbook;
    assert_eq!(replay.active_sheet().get_raw(3, 1), "North");
    assert_eq!(replay.active_sheet().get_display(7, 3), "0");
    assert_eq!(
        replay.active_sheet().table_view_spec(),
        base.active_sheet().table_view_spec()
    );
}

#[test]
fn calculated_formulas_become_visible_record_overrides_without_rewriting_the_rule() {
    let mut wb = fixture(true);
    let id = wb.active_sheet().tables()[0].id;
    wb.set_calculated_column(id, 3, 3, "=[@Amount]*2", true)
        .unwrap();
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "*2").unwrap();
    let (writes, count) = plan::replacements(&wb, 0, &view, None, None, &hits, "*3").unwrap();
    assert_eq!(count, 3);
    let mut after = prepare_table_writes(&wb, 0, &writes).unwrap();
    assert_eq!(after.active_sheet().get_display(3, 3), "90");
    assert_eq!(after.active_sheet().get_raw(4, 3), "=[@Amount]*2");
    assert_eq!(
        after.table(id).unwrap().1.columns[2].formula.as_deref(),
        Some("=[@Amount]*2")
    );
    assert!(after.active_sheet().is_calculated_exception(3, 3));
    assert!(!after.active_sheet().is_calculated_exception(4, 3));
    let commit = TableCellsCommit::capture(
        wb.active_sheet(),
        after.active_sheet(),
        writes.iter().map(|w| (w.row, w.col)),
    );
    commit.replay(&mut after, true).unwrap();
    assert!(!after.active_sheet().is_calculated_exception(3, 3));
    commit.replay(&mut after, false).unwrap();
    let raw = visigrid_io::json::export_workbook(&after, &[], 0).unwrap();
    let loaded = visigrid_io::json::import_any(&raw).unwrap().0;
    assert_eq!(loaded.active_sheet().get_display(3, 3), "90");
    assert!(loaded.active_sheet().is_calculated_exception(3, 3));
}

#[test]
fn replacing_a_sort_key_moves_the_original_record_and_preserves_hidden_values() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 3, 2, "=30");
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "30").unwrap();
    let (writes, count) = plan::replacements(&wb, 0, &view, None, None, &hits, "5").unwrap();
    assert_eq!(count, 1);
    let after = prepare_table_writes(&wb, 0, &writes).unwrap();
    assert_eq!(after.active_sheet().get_raw(3, 2), "=5");
    assert_eq!(after.active_sheet().get_raw(4, 2), "10");
    assert_eq!(rows(&after, 0).data_to_view(3), Some(3));
}

#[test]
fn stale_sources_types_sheets_and_visibility_refuse_before_any_write() {
    let wb = fixture(true);
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "West").unwrap();
    let mut changed = wb.clone();
    changed.set_cell_value_tracked(0, 6, 1, "West updated");
    assert!(plan::replacements(&changed, 0, &rows(&changed, 0), None, None, &hits, "East").is_err());
    assert_eq!(changed.active_sheet().get_raw(3, 1), "West");
    let hidden = [3].into();
    assert!(plan::replacements(&wb, 0, &view, Some(&hidden), None, &hits, "East").is_err());
    let mut changed = wb.clone();
    changed.add_sheet_named("Other").unwrap();
    assert!(plan::replacements(&changed, 1, &rows(&changed, 1), None, None, &hits, "East").is_err());
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(90), 20, 5)], 0);
    wb.set_cell_text_exact_tracked(0, 0, 0, "=1");
    let hits = plan::search(wb.active_sheet(), 0, &rows(&wb, 0), None, None, "1").unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "=1");
    assert!(plan::replacements(&wb, 0, &rows(&wb, 0), None, None, &hits, "2").is_err());
    wb.active_sheet_mut().read_only_reason = Some("Recovery".into());
    assert!(plan::replacements(&wb, 0, &rows(&wb, 0), None, None, &hits, "2")
        .unwrap_err()
        .contains("Read-only"));
}

#[test]
fn headers_footers_and_unsafe_cross_sheet_spills_refuse_the_entire_batch() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 0, 0, "Group");
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "Group").unwrap();
    assert!(plan::replacements(&wb, 0, &view, None, None, &hits, "Category")
        .unwrap_err()
        .contains("header"));
    let id = wb.active_sheet().tables()[0].id;
    wb.set_table_totals_visible(id, true, Default::default())
        .unwrap();
    wb.set_cell_value_tracked(0, 0, 0, "Total");
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "Total").unwrap();
    let (writes, _) = plan::replacements(&wb, 0, &view, None, None, &hits, "Footer").unwrap();
    assert!(prepare_table_writes(&wb, 0, &writes).is_err());
    assert_eq!(wb.active_sheet().get_raw(0, 0), "Total");
    let mut wb = fixture(true);
    let other = wb.add_sheet_named("Controls").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=1");
    wb.set_cell_text_exact_tracked(other, 1, 0, "1");
    wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(Controls!A1)");
    assert!(wb.sheet(0).unwrap().build_saved_table_view(30).is_ok());
    let other_rows = rows(&wb, other);
    let hits = plan::search(wb.sheet(other).unwrap(), other, &other_rows, None, None, "1").unwrap();
    let (writes, _) = plan::replacements(&wb, other, &other_rows, None, None, &hits, "5").unwrap();
    assert!(prepare_table_writes(&wb, other, &writes).is_err());
    assert_eq!(wb.sheet(other).unwrap().get_raw(0, 0), "=1");
    assert_eq!(wb.sheet(other).unwrap().get_raw(1, 0), "1");
}

#[test]
fn unicode_offsets_literal_text_and_whitespace_survive_replacement_and_history() {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(90), 20, 5)], 0);
    wb.set_cell_text_exact_tracked(0, 0, 0, "  \u{212a}K \u{130}i  ");
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "k").unwrap();
    assert_eq!(
        hits.iter().map(|h| (h.start, h.end)).collect::<Vec<_>>(),
        [(2, 5), (5, 6)]
    );
    let (writes, _) = plan::replacements(&wb, 0, &view, None, None, &hits, "Z").unwrap();
    let after = prepare_table_writes(&wb, 0, &writes).unwrap();
    assert_eq!(after.active_sheet().get_raw(0, 0), "  ZZ \u{130}i  ");
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "i").unwrap();
    let (writes, _) = plan::replacements(&wb, 0, &view, None, None, &hits, "\u{e9}").unwrap();
    assert_eq!(
        prepare_table_writes(&wb, 0, &writes)
            .unwrap()
            .active_sheet()
            .get_raw(0, 0),
        "  \u{212a}K \u{e9}\u{e9}  "
    );
    wb.set_cell_text_exact_tracked(0, 0, 0, "00123");
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "123").unwrap();
    let (writes, _) = plan::replacements(&wb, 0, &view, None, None, &hits, "456").unwrap();
    let after = prepare_table_writes(&wb, 0, &writes).unwrap();
    assert!(matches!(
        after.active_sheet().get_cell(0, 0).value(),
        ValueRef::Text("00456")
    ));
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "00123").unwrap();
    let (writes, _) = plan::replacements(&wb, 0, &view, None, None, &hits, "=1+1").unwrap();
    let mut after = prepare_table_writes(&wb, 0, &writes).unwrap();
    assert!(matches!(
        after.active_sheet().get_cell(0, 0).value(),
        ValueRef::Text("=1+1")
    ));
    let commit = TableCellsCommit::capture(wb.active_sheet(), after.active_sheet(), [(0, 0)]);
    commit.replay(&mut after, true).unwrap();
    assert!(matches!(
        after.active_sheet().get_cell(0, 0).value(),
        ValueRef::Text("00123")
    ));
}

#[test]
fn reference_matches_respect_tokens_quotes_and_structured_columns_and_replace_one_occurrence() {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(90), 20, 5)], 0);
    let source = "=A1+A10+$A$1+LEN(\"A1 \"\"A1\"\"\")+'A1'!A1+SUM(Sales[A1])+A1!B1";
    wb.set_cell_value_tracked(0, 10, 0, source);
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "A1").unwrap();
    assert_eq!(hits.len(), 2);
    let (writes, count) = plan::replacements(&wb, 0, &view, None, None, &hits[..1], "B2").unwrap();
    assert_eq!(count, 1);
    assert_eq!(
        writes[0].value.as_deref(),
        Some("=B2+A10+$A$1+LEN(\"A1 \"\"A1\"\"\")+'A1'!A1+SUM(Sales[A1])+A1!B1")
    );
    let absolute = plan::search(wb.active_sheet(), 0, &view, None, None, "$A$1").unwrap();
    assert_eq!(absolute.len(), 1);
    assert_eq!(absolute[0].kind, Some(MatchKind::Formula));
}

#[test]
fn display_matches_are_read_only_noops_are_empty_and_excess_matches_refuse_without_truncation() {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(90), 20, 5)], 0);
    wb.set_cell_value_tracked(0, 0, 0, "12345");
    wb.set_cell_text_exact_tracked(0, 1, 0, "$12,345%");
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "12345").unwrap();
    assert_eq!(hits.len(), 2);
    assert!(hits[0].kind.is_none());
    let (writes, count) = plan::replacements(&wb, 0, &view, None, None, &hits, "0").unwrap();
    assert_eq!(count, 1);
    assert_eq!(writes[0].value.as_deref(), Some("$0%"));
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "$12,345%").unwrap();
    // Normalization only selects the numeric portion; replacing it with itself is a no-op.
    let (writes, count) = plan::replacements(&wb, 0, &view, None, None, &hits, "12,345").unwrap();
    assert!(writes.is_empty());
    assert_eq!(count, 0);
    wb.set_cell_text_exact_tracked(0, 2, 0, &"x".repeat(100_001));
    assert!(plan::search(wb.active_sheet(), 0, &view, None, None, "x")
        .unwrap_err()
        .contains("100,000"));
}

#[test]
fn replacement_on_an_unfiltered_sheet_recalculates_other_sheet_views_and_replays() {
    let mut wb = fixture(true);
    let other = wb.add_sheet_named("Controls").unwrap();
    wb.set_cell_value_tracked(other, 0, 0, "=1");
    wb.set_cell_value_tracked(0, 3, 1, "=IF(Controls!A1=1,\"West\",\"East\")");
    assert!(rows(&wb, 0).data_to_view(3).is_some());
    let view = rows(&wb, other);
    let hits = plan::search(wb.sheet(other).unwrap(), other, &view, None, None, "1").unwrap();
    let (writes, _) = plan::replacements(&wb, other, &view, None, None, &hits, "2").unwrap();
    let mut after = prepare_table_writes(&wb, other, &writes).unwrap();
    assert!(rows(&after, 0).data_to_view(3).is_none());
    assert_eq!(after.sheet(0).unwrap().get_raw(4, 1), "East");
    let commit = TableCellsCommit::capture(
        wb.sheet(other).unwrap(),
        after.sheet(other).unwrap(),
        [(0, 0)],
    );
    commit.replay(&mut after, true).unwrap();
    assert!(rows(&after, 0).data_to_view(3).is_some());
    commit.replay(&mut after, false).unwrap();
    assert!(rows(&after, 0).data_to_view(3).is_none());
}

#[test]
fn ordinary_filtered_rows_and_merged_titles_use_the_same_visible_cell_contract() {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(90), 20, 5)], 0);
    for row in 0..3 {
        wb.set_cell_text_exact_tracked(0, row, 0, "Find me");
    }
    let mut view = RowView::new(20);
    let mut mask = vec![true; 20];
    mask[1] = false;
    view.apply_filter(mask);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "me").unwrap();
    assert_eq!(hits.iter().map(|h| h.row).collect::<Vec<_>>(), [0, 2]);
    let (writes, _) = plan::replacements(&wb, 0, &view, None, None, &hits, "you").unwrap();
    let after = prepare_table_writes(&wb, 0, &writes).unwrap();
    assert_eq!(after.active_sheet().get_raw(1, 0), "Find me");
    let mut wb = fixture(true);
    wb.set_cell_text_exact_tracked(0, 0, 0, "Title");
    wb.active_sheet_mut()
        .merged_regions
        .push(visigrid_engine::sheet::MergedRegion {
            start: (0, 0),
            end: (0, 1),
        });
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "Title").unwrap();
    let (writes, _) = plan::replacements(&wb, 0, &view, None, None, &hits, "New title").unwrap();
    let mut after = prepare_table_writes(&wb, 0, &writes).unwrap();
    let commit = TableCellsCommit::capture(wb.active_sheet(), after.active_sheet(), [(0, 0)]);
    commit.replay(&mut after, true).unwrap();
    assert_eq!(after.active_sheet().get_raw(0, 0), "Title");
    assert_eq!(
        after.active_sheet().merged_regions,
        wb.active_sheet().merged_regions
    );
}

#[test]
fn repetitive_text_and_lowercase_expansions_keep_nonoverlapping_original_spans() {
    let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(90), 20, 5)], 0);
    wb.set_cell_text_exact_tracked(0, 0, 0, &("a".repeat(100_000) + "b"));
    let query = "a".repeat(10_000) + "b";
    let view = rows(&wb, 0);
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, &query).unwrap();
    assert_eq!(hits.len(), 1);
    assert_eq!((hits[0].start, hits[0].end), (90_000, 100_001));
    wb.set_cell_text_exact_tracked(0, 0, 0, "aaaaa");
    let hits = plan::search(wb.active_sheet(), 0, &view, None, None, "aa").unwrap();
    assert_eq!(
        hits.iter().map(|h| (h.start, h.end)).collect::<Vec<_>>(),
        [(0, 2), (2, 4)]
    );
    let (writes, count) = plan::replacements(&wb, 0, &view, None, None, &hits, "z").unwrap();
    assert_eq!(count, 2);
    assert_eq!(writes[0].value.as_deref(), Some("zza"));
}

#[test]
fn hidden_columns_are_excluded_and_stale_hits_cannot_be_replaced() {
    let mut wb = fixture(true);
    wb.set_cell_value_tracked(0, 3, 2, "West");
    let view = rows(&wb, 0);
    let all = plan::search(wb.active_sheet(), 0, &view, None, None, "West").unwrap();
    let hidden = [1].into();
    let hits = plan::search(wb.active_sheet(), 0, &view, None, Some(&hidden), "West").unwrap();
    assert_eq!(hits.iter().map(|h| (h.row, h.col)).collect::<Vec<_>>(), [(3, 2)]);
    assert!(plan::replacements(&wb, 0, &view, None, Some(&hidden), &all, "North").is_err());
    let (writes, count) = plan::replacements(&wb, 0, &view, None, Some(&hidden), &hits, "North").unwrap();
    assert_eq!(count, 1);
    let after = prepare_table_writes(&wb, 0, &writes).unwrap();
    assert_eq!(after.active_sheet().get_raw(3, 2), "North");
    assert_eq!(after.active_sheet().get_raw(3, 1), "West");
}
