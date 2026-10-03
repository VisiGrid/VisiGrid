//! Cut captures the old projection, then clears visible source records atomically.
//! Paste is a separate operation, matching the ordinary clipboard contract.
use crate::{
    app::Spreadsheet,
    clipboard::InternalClipboard,
    table_edit::{view_safe_selection_rows, TableCellWrite},
};
use gpui::{ClipboardItem, Context};
use visigrid_engine::{filter::RowView, sheet::Sheet};

type Rect = ((usize, usize), (usize, usize));

pub(crate) fn plan_cut(
    sheet: &Sheet,
    view: &RowView,
    rect: Rect,
) -> Result<(InternalClipboard, Vec<TableCellWrite>), String> {
    let ((_, c0), (_, c1)) = rect;
    let rows = view_safe_selection_rows(sheet, view, rect)?;
    let source_row = rows[0];
    let mut clipboard = InternalClipboard {
        raw_tsv: String::new(),
        raw_cells: Vec::new(),
        values: Vec::new(),
        formats: Vec::new(),
        comments: Vec::new(),
        source: (source_row, c0),
        source_rows: rows.clone(),
        source_formulas: Vec::new(),
        id: rand::random(),
        merges: Vec::new(),
        created_at: std::time::Instant::now(),
    };
    let mut writes = Vec::new();
    for row in rows {
        clipboard
            .raw_cells
            .push((c0..=c1).map(|col| sheet.get_raw(row, col)).collect());
        clipboard.values.push(
            (c0..=c1)
                .map(|col| sheet.get_computed_value(row, col))
                .collect(),
        );
        clipboard.formats.push(
            (c0..=c1)
                .map(|col| sheet.get_format(row, col).clone())
                .collect(),
        );
        clipboard.comments.push(
            (c0..=c1)
                .map(|col| sheet.comment(row, col).cloned())
                .collect(),
        );
        clipboard.source_formulas.push(
            (c0..=c1)
                .map(|col| {
                    sheet
                        .get_cell_opt(row, col)
                        .is_some_and(|cell| cell.value().is_formula())
                })
                .collect(),
        );
        for col in c0..=c1 {
            let mut write = TableCellWrite::value(row, col, String::new());
            write.comment = Some(None);
            writes.push(write);
        }
    }
    clipboard.raw_tsv = clipboard
        .raw_cells
        .iter()
        .map(|r| r.join("\t"))
        .collect::<Vec<_>>()
        .join("\n");
    Ok((clipboard, writes))
}

impl Spreadsheet {
    pub(crate) fn cut_table_view(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) {
            return;
        }
        self.sync_table_view(cx);
        // Only the primary selection participates, just like ordinary Cut.
        let (clipboard, writes) =
            match plan_cut(self.sheet(cx), &self.row_view, self.selection_range()) {
                Ok(plan) => plan,
                Err(error) => {
                    self.status_message = Some(error);
                    cx.notify();
                    return;
                }
            };
        if !self.finish_table_selection_write(Ok(writes), "Cut cells", None, cx) {
            return;
        }
        // Publish the clipboard only after the guarded candidate succeeds. A
        // rejected cut leaves the previous system/internal clipboard untouched.
        let text = if clipboard.raw_tsv.is_empty() {
            "\n".into()
        } else {
            clipboard.raw_tsv.clone()
        };
        let metadata = format!("\"{}\"", clipboard.id);
        self.internal_clipboard = Some(clipboard);
        if self.mode != crate::mode::Mode::FormatPainter {
            self.format_painter = None;
        }
        // The source may have moved or disappeared. Don't outline different records.
        self.clipboard_visual_range = None;
        cx.write_to_clipboard(ClipboardItem::new_string_with_json_metadata(text, metadata));
        cx.notify();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        table_cell_history::TableCellsCommit,
        table_edit::{prepare_table_writes, tests::fixture},
    };
    use visigrid_engine::cell::CellComment;
    fn view(sheet: &Sheet) -> RowView {
        sheet
            .build_saved_table_view(sheet.rows)
            .unwrap()
            .unwrap()
            .rows()
            .clone()
    }

    #[test]
    fn cut_captures_display_order_and_clears_only_visible_records() {
        let mut before = fixture(true);
        before.active_sheet_mut().set_comment(
            5,
            2,
            Some(CellComment {
                text: "keep me".into(),
                author: String::new(),
            }),
        );
        let mut format = before.active_sheet().get_format(5, 2).clone();
        format.bold = true;
        before.active_sheet_mut().set_format(5, 2, format.clone());
        let (ic, writes) = plan_cut(
            before.active_sheet(),
            &view(before.active_sheet()),
            ((3, 2), (6, 3)),
        )
        .unwrap();
        assert_eq!(ic.source_rows, vec![5, 3, 6]);
        assert_eq!(ic.raw_tsv, "20\t=C6*2\n30\t=C4*2\n40\t=C7*2");
        assert!(ic.formats[0][0].bold);
        assert!(ic.comments[0][0].is_some());
        let mut after = prepare_table_writes(&before, 0, &writes).unwrap();
        assert_eq!(after.active_sheet().get_raw(4, 2), "10");
        assert_eq!(after.active_sheet().get_raw(4, 3), "=C5*2");
        assert_eq!(after.active_sheet().get_raw(0, 1), "=SUM(C4:C7)");
        assert_eq!(after.active_sheet().get_display(0, 1), "10");
        for row in [5, 3, 6] {
            assert_eq!(after.active_sheet().get_raw(row, 2), "");
            assert_eq!(after.active_sheet().get_raw(row, 3), "");
        }
        assert_eq!(after.active_sheet().get_format(5, 2), format);
        assert!(after.active_sheet().comment(5, 2).is_none());
        let commit = TableCellsCommit::capture(
            before.active_sheet(),
            after.active_sheet(),
            writes.iter().map(|w| (w.row, w.col)),
        );
        assert_eq!(commit.patches.len(), 6);
        commit.replay(&mut after, true).unwrap();
        assert_eq!(after.active_sheet().get_display(0, 1), "100");
        assert_eq!(
            after.active_sheet().comment(5, 2),
            before.active_sheet().comment(5, 2)
        );
        commit.replay(&mut after, false).unwrap();
        assert_eq!(after.active_sheet().get_display(0, 1), "10");
    }

    #[test]
    fn cutting_filter_keys_can_hide_every_record_and_undo_restores_them() {
        let before = fixture(true);
        let (ic, writes) = plan_cut(
            before.active_sheet(),
            &view(before.active_sheet()),
            ((3, 1), (6, 1)),
        )
        .unwrap();
        let mut after = prepare_table_writes(&before, 0, &writes).unwrap();
        let view = after
            .active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap();
        assert!(view.visible_body_rows(4, 1).is_err());
        assert_eq!(after.active_sheet().get_raw(4, 1), "East");
        assert_eq!(ic.raw_tsv, "West\nWest\nWest");
        let commit = TableCellsCommit::capture(
            before.active_sheet(),
            after.active_sheet(),
            writes.iter().map(|w| (w.row, w.col)),
        );
        commit.replay(&mut after, true).unwrap();
        assert_eq!(
            after
                .active_sheet()
                .build_saved_table_view(30)
                .unwrap()
                .unwrap()
                .visible_body_rows(4, 3)
                .unwrap(),
            vec![5, 3, 6]
        );
        commit.replay(&mut after, false).unwrap();
    }

    #[test]
    fn invalid_source_or_recalculation_rejects_entire_cut() {
        let mut wb = fixture(false);
        for rect in [
            ((2, 1), (3, 2)),
            ((3, 0), (4, 2)),
            ((3, 1), (7, 2)),
            ((3, 1), (4, 4)),
        ] {
            assert!(plan_cut(wb.active_sheet(), &view(wb.active_sheet()), rect).is_err());
        }
        // Clearing C4 would expand an adjacent array through the Table body.
        wb.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(31-C4)");
        let (_, writes) = plan_cut(
            wb.active_sheet(),
            &view(wb.active_sheet()),
            ((3, 2), (6, 2)),
        )
        .unwrap();
        let revision = wb.revision();
        assert!(prepare_table_writes(&wb, 0, &writes).is_err());
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.active_sheet().get_raw(3, 2), "30");
        assert_eq!(wb.active_sheet().get_raw(5, 2), "20");
    }

    #[test]
    fn clipboard_keeps_multiline_text_and_comment_only_trailing_cells() {
        let mut wb = fixture(true);
        wb.set_cell_text_tracked(0, 5, 3, "001\t=hello\nworld");
        wb.set_cell_value_tracked(0, 6, 3, "");
        wb.active_sheet_mut().set_comment(
            6,
            3,
            Some(CellComment {
                text: "note".into(),
                author: String::new(),
            }),
        );
        let (ic, writes) = plan_cut(
            wb.active_sheet(),
            &view(wb.active_sheet()),
            ((4, 3), (6, 3)),
        )
        .unwrap();
        assert_eq!(
            ic.raw_cells,
            vec![vec!["001\t=hello\nworld"], vec!["=C4*2"], vec![""]]
        );
        assert_eq!(
            ic.source_formulas,
            vec![vec![false], vec![true], vec![false]]
        );
        assert!(ic.comments[2][0].is_some());
        let after = prepare_table_writes(&wb, 0, &writes).unwrap();
        assert!(after.active_sheet().comment(6, 3).is_none());
    }
    #[test]
    fn cut_then_paste_rebases_formulas_and_retains_text_comments_and_separate_history() {
        use crate::clipboard::{table_paste_writes, TablePasteKind};
        let mut before = fixture(true);
        before.set_cell_text_tracked(0, 5, 3, "001\t=hello\nworld");
        before.set_cell_text_tracked(0, 6, 3, "=literal");
        before.active_sheet_mut().set_comment(
            6,
            3,
            Some(CellComment {
                text: "note".into(),
                author: "A".into(),
            }),
        );
        let (ic, writes) = plan_cut(
            before.active_sheet(),
            &view(before.active_sheet()),
            ((4, 3), (6, 3)),
        )
        .unwrap();
        let cut = prepare_table_writes(&before, 0, &writes).unwrap();
        let cut_commit = TableCellsCommit::capture(
            before.active_sheet(),
            cut.active_sheet(),
            writes.iter().map(|w| (w.row, w.col)),
        );
        // Paste into a different canonical order; the source snapshot is unchanged.
        let paste = table_paste_writes(
            &ic.raw_cells,
            Some(&ic),
            TablePasteKind::Contents,
            vec![(6, 3, 0, 0), (5, 3, 1, 0), (3, 3, 2, 0)],
        );
        let mut after = prepare_table_writes(&cut, 0, &paste).unwrap();
        assert_eq!(after.active_sheet().get_raw(6, 3), "001\t=hello\nworld");
        assert_eq!(after.active_sheet().get_raw(5, 3), "=C6*2");
        assert_eq!(after.active_sheet().get_display(5, 3), "40");
        assert_eq!(after.active_sheet().get_raw(3, 3), "=literal");
        assert!(!after.active_sheet().get_cell(3, 3).value().is_formula());
        assert_eq!(
            after.active_sheet().comment(3, 3),
            before.active_sheet().comment(6, 3)
        );
        let paste_commit = TableCellsCommit::capture(
            cut.active_sheet(),
            after.active_sheet(),
            paste.iter().map(|w| (w.row, w.col)),
        );
        paste_commit.replay(&mut after, true).unwrap();
        for row in [3, 5, 6] {
            assert_eq!(after.active_sheet().get_raw(row, 3), "");
        }
        cut_commit.replay(&mut after, true).unwrap();
        for row in [3, 4, 5, 6] {
            assert_eq!(
                after.active_sheet().get_raw(row, 3),
                before.active_sheet().get_raw(row, 3)
            );
        }
        cut_commit.replay(&mut after, false).unwrap();
        paste_commit.replay(&mut after, false).unwrap();
        assert_eq!(after.active_sheet().get_raw(6, 3), "001\t=hello\nworld");
    }
}
