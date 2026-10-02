//! Header paste is a schema rename, never a batch of ordinary cell writes.
use crate::{
    app::Spreadsheet,
    clipboard::{InternalClipboard, TablePasteKind},
};
use gpui::Context;
use visigrid_engine::workbook::{TableCommit, Workbook};

/// Preserve internal cell boundaries and literal external names. Headers cannot
/// contain newlines, so reject multiline input before CSV parsing can skip rows.
fn header_grid(
    text: &str,
    internal: Option<&InternalClipboard>,
    values: bool,
) -> Result<Vec<Vec<String>>, String> {
    if let Some(ic) = internal {
        if !ic.merges.is_empty() {
            return Err(
                "Cannot paste merged cells into Table headers. Unmerge the source first.".into(),
            );
        }
        return Ok(if values {
            ic.values
                .iter()
                .map(|row| {
                    row.iter()
                        .map(Spreadsheet::value_to_canonical_string)
                        .collect()
                })
                .collect()
        } else {
            ic.raw_cells.clone()
        });
    }
    let text = text
        .strip_suffix("\r\n")
        .or_else(|| text.strip_suffix('\n'))
        .unwrap_or(text);
    if text.contains(['\n', '\r']) {
        return Err(
            "Paste a single row of names into the Table headers. Nothing was pasted.".into(),
        );
    }
    if text.is_empty() {
        return Ok(vec![vec![String::new()]]);
    }
    // Match ordinary paste's thousands-separator guard, without coercing names.
    if !text.contains('\t') && visigrid_engine::cell::try_parse_number(text.trim()).is_some() {
        return Ok(vec![vec![text.to_string()]]);
    }
    let grid = visigrid_io::csv::parse_delimited_text(text);
    if grid.is_empty() {
        return Err("The clipboard has no header names.".into());
    }
    Ok(grid)
}

fn validate_views(wb: &Workbook) -> Result<(), String> {
    for sheet in wb.sheets() {
        sheet.build_saved_table_view(crate::app::NUM_ROWS.min(sheet.rows))?;
    }
    Ok(())
}

/// A name-only schema change may pass the saved-view history gate. Formula
/// rules can also change because references to renamed columns are rewritten.
pub(crate) fn is_header_rename(commit: &TableCommit) -> bool {
    let (Some(before), Some(after)) = (commit.before_table(), commit.after_table()) else {
        return false;
    };
    if before.columns.len() != after.columns.len() || commit.inserted_header_row().is_some() {
        return false;
    }
    let changed = before
        .columns
        .iter()
        .zip(&after.columns)
        .any(|(a, b)| a.name != b.name);
    let mut normalized = after.clone();
    for (a, b) in before.columns.iter().zip(&mut normalized.columns) {
        b.name = a.name.clone();
        b.formula = a.formula.clone();
    }
    changed && *before == normalized
}

/// Validate the whole destination and schema before publishing any changes.
fn prepare_header_paste(
    wb: &Workbook,
    sheet_index: usize,
    row: usize,
    col: usize,
    grid: &[Vec<String>],
) -> Result<Option<(Workbook, TableCommit)>, String> {
    let sheet = wb.sheet(sheet_index).ok_or("The sheet no longer exists.")?;
    let table = sheet
        .table_header_at(row, col)
        .ok_or("Select a Table header before pasting names.")?;
    if grid.len() != 1 || grid[0].is_empty() {
        return Err(
            "Paste a single row of names into the Table headers. Nothing was pasted.".into(),
        );
    }
    if col
        .checked_add(grid[0].len())
        .is_none_or(|end| end > table.range.end_col + 1)
    {
        return Err("The header paste extends beyond this Table. Nothing was pasted.".into());
    }
    let mut names: Vec<_> = table.columns.iter().map(|c| c.name.clone()).collect();
    let offset = col - table.range.start_col;
    names[offset..offset + grid[0].len()].clone_from_slice(&grid[0]);
    let mut used = std::collections::HashSet::new();
    for name in &names {
        if name.starts_with('=') {
            return Err("Table headers must be names. Use Paste Values to use copied formula results as names.".into());
        }
        visigrid_engine::table::validate_column_name(name)?;
        if !used.insert(name.to_lowercase()) {
            return Err(format!("The name {name:?} is already used by another column in {}. Each header needs a unique name.", table.name));
        }
    }

    if names
        .iter()
        .zip(&table.columns)
        .all(|(name, c)| *name == c.name)
    {
        return Ok(None);
    }
    let mut candidate = wb.clone();
    let commit = candidate.rename_table_columns(table.id, &names)?;
    if let Some(error) = candidate.take_incremental_errors().first() {
        return Err(format!(
            "The header rename could not be recalculated: {error:?}"
        ));
    }
    validate_views(&candidate)?;
    Ok(Some((candidate, commit)))
}

pub(crate) fn prepare_header_replay(
    wb: &Workbook,
    commit: &TableCommit,
    undo: bool,
) -> Result<Workbook, String> {
    if !is_header_rename(commit) {
        return Err("This history entry is not a header rename.".into());
    }
    let mut candidate = wb.clone();
    candidate.apply_table_commit(commit, undo)?;
    if let Some(error) = candidate.take_incremental_errors().first() {
        return Err(format!(
            "The header rename could not be recalculated: {error:?}"
        ));
    }
    validate_views(&candidate)?;
    Ok(candidate)
}

impl Spreadsheet {
    /// Return true when the destination is a header, including a refused paste.
    pub(crate) fn paste_table_headers(
        &mut self,
        kind: TablePasteKind,
        cx: &mut Context<Self>,
    ) -> bool {
        if self.mode.is_editing() {
            return false;
        }
        let (view_row, col) = self.view_state.selected;
        let row = self.row_view.view_to_data(view_row);
        let Some(table) = self.sheet(cx).table_header_at(row, col) else {
            return false;
        };
        let range = table.range;
        let name = table.name.clone();
        let result = (|| {
            if !matches!(kind, TablePasteKind::Contents | TablePasteKind::Values) {
                return Err("Use Paste or Paste Values to rename Table headers. Header formatting and comments are retained.".to_string());
            }
            if !self.view_state.additional_selections.is_empty() {
                return Err(
                    "Select one contiguous row of Table headers before pasting names.".into(),
                );
            }
            let ((r0, c0), (r1, c1)) = self.selection_range();
            if r0 != r1 || c0 < range.start_col || c1 > range.end_col {
                return Err("Select only headers within one Table before pasting names.".into());
            }
            let item = cx.read_from_clipboard();
            let text = item.as_ref().and_then(|i| i.text());
            let internal = Self::is_internal_paste(
                self.internal_clipboard.as_ref(),
                text.as_deref(),
                item.as_ref().and_then(|i| i.metadata()).map(|s| s.as_str()),
            );
            let ic = internal
                .then_some(self.internal_clipboard.as_ref())
                .flatten();
            let text = ic
                .map(|ic| ic.raw_tsv.as_str())
                .or(text.as_deref())
                .ok_or("The clipboard is empty.")?;
            let grid = header_grid(text, ic, kind == TablePasteKind::Values)?;
            if grid.len() == 1 && grid[0].len() == 1 && c0 != c1 {
                return Err(
                    "Paste one name per header; a single name cannot fill multiple headers.".into(),
                );
            }
            let prepared =
                prepare_header_paste(self.wb(cx), self.sheet_index(cx), row, col, &grid)?;
            if let Some((candidate, commit)) = prepared {
                self.validate_saved_view_layout(&candidate)?;
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                self.record_table_commit(
                    commit,
                    format!(
                        "Rename {} header{}: {name}",
                        grid[0].len(),
                        if grid[0].len() == 1 { "" } else { "s" }
                    ),
                    cx,
                );
                self.clipboard_visual_range = None;
            } else {
                self.status_message = Some("Table header names are unchanged.".into());
                cx.notify();
            }
            Ok(())
        })();
        if let Err(error) = result {
            self.status_message = Some(error);
            cx.notify();
        }
        true
    }

    pub(crate) fn replay_table_headers(
        &mut self,
        commit: &TableCommit,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let result = prepare_header_replay(self.wb(cx), commit, undo).and_then(|candidate| {
            self.validate_saved_view_layout(&candidate)?;
            Ok(candidate)
        });
        match result {
            Ok(candidate) => {
                self.workbook
                    .update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                self.bump_cells_rev();
                self.is_modified = true;
                cx.notify();
                true
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                false
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        history::{History, UndoAction},
        table_edit::tests::fixture,
    };
    use visigrid_engine::{
        cell::{CellComment, CellFormat},
        formula::eval::Value,
        sheet::SheetId,
    };

    fn grid(names: &[&str]) -> Vec<Vec<String>> {
        vec![names.iter().map(|s| s.to_string()).collect()]
    }

    fn projection(wb: &Workbook) -> Vec<usize> {
        wb.active_sheet()
            .build_saved_table_view(30)
            .unwrap()
            .unwrap()
            .visible_body_rows(4, 3)
            .unwrap()
    }

    #[test]
    fn swaps_preserve_column_identity_rules_overrides_filters_and_cross_sheet_formulas() {
        let mut before = fixture(true);
        let id = before.active_sheet().tables()[0].id;
        before
            .set_calculated_column(id, 3, 3, "=[@Amount]*2", true)
            .unwrap();
        before.set_cell_value_tracked(0, 5, 3, "999");
        let other = before.add_sheet();
        before.set_cell_value_tracked(other, 0, 0, "=SUM(Sales[Amount])");
        before.active_sheet_mut().set_comment(
            2,
            1,
            Some(CellComment {
                text: "keep".into(),
                author: "me".into(),
            }),
        );
        let format = CellFormat {
            bold: true,
            ..Default::default()
        };
        before.active_sheet_mut().set_format(2, 1, format.clone());
        let (after, commit) = prepare_header_paste(&before, 0, 2, 1, &grid(&["Amount", "Group"]))
            .unwrap()
            .unwrap();
        assert!(is_header_rename(&commit));
        assert_eq!(
            after.sheet(other).unwrap().get_raw(0, 0),
            "=SUM(Sales[Group])"
        );
        assert_eq!(after.sheet(other).unwrap().get_display(0, 0), "100");
        assert_eq!(after.active_sheet().get_raw(3, 3), "=[@[Group]]*2");
        assert_eq!(after.active_sheet().get_raw(5, 3), "999");
        assert!(after.active_sheet().is_calculated_exception(5, 3));
        assert_eq!(
            after.table(id).unwrap().1.columns[2].formula.as_deref(),
            Some("=[@[Group]]*2")
        );
        assert_eq!(
            after.active_sheet().table_view_spec(),
            before.active_sheet().table_view_spec()
        );
        assert_eq!(projection(&after), projection(&before));
        assert_eq!(after.active_sheet().get_format(2, 1), format);
        assert_eq!(
            after.active_sheet().comment(2, 1),
            before.active_sheet().comment(2, 1)
        );
        assert_eq!(
            before
                .table(id)
                .unwrap()
                .1
                .columns
                .iter()
                .map(|c| c.id)
                .collect::<Vec<_>>(),
            after
                .table(id)
                .unwrap()
                .1
                .columns
                .iter()
                .map(|c| c.id)
                .collect::<Vec<_>>()
        );
        let undone = prepare_header_replay(&after, &commit, true).unwrap();
        assert_eq!(undone.table(id), before.table(id));
        assert_eq!(
            undone.sheet(other).unwrap().get_raw(0, 0),
            "=SUM(Sales[Amount])"
        );
        let redone = prepare_header_replay(&undone, &commit, false).unwrap();
        assert_eq!(redone.table(id), after.table(id));
        assert_eq!(projection(&redone), projection(&before));
    }

    #[test]
    fn partial_paste_escapes_names_and_retains_untouched_headers() {
        let mut before = fixture(false);
        before.set_cell_value_tracked(0, 0, 0, "=SUM(Sales[Amount])+SUM(Sales[Result])");
        let (after, _) = prepare_header_paste(&before, 0, 2, 2, &grid(&["金額 [net]", "O'Brien"]))
            .unwrap()
            .unwrap();
        assert_eq!(after.active_sheet().get_raw(2, 1), "Group");
        assert_eq!(after.active_sheet().get_raw(2, 2), "金額 [net]");
        assert_eq!(
            after.active_sheet().get_raw(0, 0),
            "=SUM(Sales[金額 '[net']])+SUM(Sales[O''Brien])"
        );
        assert_eq!(after.active_sheet().get_display(0, 0), "300");
        assert_eq!(projection(&after), projection(&before));
    }

    #[test]
    fn invalid_names_shape_and_bounds_leave_every_cell_and_revision_unchanged() {
        let before = fixture(true);
        let revision = before.revision();
        let schema = before.active_sheet().tables()[0].clone();
        let duplicate =
            prepare_header_paste(&before, 0, 2, 1, &grid(&["Same", "same"])).unwrap_err();
        assert!(duplicate.contains("already used"));
        assert!(!duplicate.contains("IDs"));

        let cells: Vec<_> = before
            .active_sheet()
            .cells_iter()
            .map(|(p, c)| (p, c.raw_display()))
            .collect();
        for invalid in [
            grid(&["Same", "same"]),
            grid(&["Result"]),
            grid(&[""]),
            grid(&[" padded"]),
            grid(&["=1+1"]),
            grid(&["a\tb"]),
            grid(&["a\nb"]),
            grid(&["A", "B", "C", "D"]),
            vec![vec!["A".into()], vec!["B".into()]],
            vec![],
        ] {
            assert!(
                prepare_header_paste(&before, 0, 2, 1, &invalid).is_err(),
                "{invalid:?}"
            );
            assert_eq!(before.revision(), revision);
            assert_eq!(before.active_sheet().tables()[0], schema);
            assert_eq!(
                before
                    .active_sheet()
                    .cells_iter()
                    .map(|(p, c)| (p, c.raw_display()))
                    .collect::<Vec<_>>(),
                cells
            );
        }
        assert!(prepare_header_paste(&before, 0, 3, 1, &grid(&["A"])).is_err());
        assert!(prepare_header_paste(&before, 0, 2, 0, &grid(&["A", "B"])).is_err());
    }

    #[test]
    fn unchanged_paste_has_no_commit() {
        let before = fixture(true);
        assert!(
            prepare_header_paste(&before, 0, 2, 1, &grid(&["Group", "Amount", "Result"]))
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn external_headers_parse_quotes_crlf_and_literal_numbers_without_dropping_rows() {
        assert_eq!(
            header_grid("\"North, South\"\t金額\r\n", None, false).unwrap(),
            grid(&["North, South", "金額"])
        );
        assert_eq!(
            header_grid("\"North, South\",Amount", None, false).unwrap(),
            grid(&["North, South", "Amount"])
        );
        assert_eq!(
            header_grid("001\t2026", None, true).unwrap(),
            grid(&["001", "2026"])
        );
        assert_eq!(header_grid("1,000", None, false).unwrap(), grid(&["1,000"]));
        assert_eq!(
            header_grid("A\t\r\n", None, false).unwrap(),
            grid(&["A", ""])
        );
        for text in ["A\nB", "A\n\n", "\"A\nB\"\tC", "A\rB"] {
            assert!(header_grid(text, None, false).is_err());
        }
    }

    fn clipboard() -> InternalClipboard {
        InternalClipboard {
            raw_tsv: "=\"Area\"\t=\"Revenue\"".into(),
            raw_cells: grid(&["=\"Area\"", "=\"Revenue\""]),
            values: vec![vec![
                Value::Text("Area".into()),
                Value::Text("Revenue".into()),
            ]],
            formats: vec![vec![CellFormat::default(); 2]],
            comments: vec![vec![None; 2]],
            source: (10, 0),
            source_rows: vec![10],
            source_formulas: vec![vec![true; 2]],
            id: 17,
            merges: vec![],
            created_at: std::time::Instant::now(),
        }
    }

    #[test]
    fn internal_values_use_formula_results_and_keep_exact_cell_boundaries() {
        let before = fixture(true);
        let mut ic = clipboard();
        let raw = header_grid(&ic.raw_tsv, Some(&ic), false).unwrap();
        assert!(prepare_header_paste(&before, 0, 2, 1, &raw).is_err());
        let values = header_grid(&ic.raw_tsv, Some(&ic), true).unwrap();
        let (after, _) = prepare_header_paste(&before, 0, 2, 1, &values)
            .unwrap()
            .unwrap();
        assert_eq!(after.active_sheet().get_raw(2, 2), "Revenue");
        ic.raw_cells = grid(&["A\tB"]);
        assert_eq!(
            header_grid(&ic.raw_tsv, Some(&ic), false).unwrap(),
            grid(&["A\tB"])
        );
        ic.raw_cells = vec![vec!["A".into()], vec![String::new()]];
        assert!(prepare_header_paste(
            &before,
            0,
            2,
            1,
            &header_grid(&ic.raw_tsv, Some(&ic), false).unwrap()
        )
        .is_err());
    }

    #[test]
    fn replay_rejects_new_dependents_without_rewriting_the_existing_workbook() {
        let before = fixture(true);
        let (mut after, commit) =
            prepare_header_paste(&before, 0, 2, 1, &grid(&["Area", "Revenue"]))
                .unwrap()
                .unwrap();
        after.set_cell_value_tracked(0, 0, 0, "=SUM(Sales[Revenue])");
        let revision = after.revision();
        assert!(prepare_header_replay(&after, &commit, true).is_err());
        assert_eq!(after.revision(), revision);
        assert_eq!(after.active_sheet().get_raw(2, 1), "Area");
        assert_eq!(after.active_sheet().get_raw(0, 0), "=SUM(Sales[Revenue])");
    }

    #[test]
    fn recalculation_that_invalidates_a_view_rejects_paste_and_undo() {
        let mut before = fixture(true);
        before.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(IF(B3=\"Group\",1,5))");
        assert!(before.active_sheet().build_saved_table_view(30).is_ok());
        let revision = before.revision();
        assert!(prepare_header_paste(&before, 0, 2, 1, &grid(&["Area"])).is_err());
        assert_eq!(before.revision(), revision);
        assert_eq!(before.active_sheet().get_raw(2, 1), "Group");
        before.set_cell_value_tracked(0, 0, 0, "");
        let (mut after, commit) = prepare_header_paste(&before, 0, 2, 1, &grid(&["Area"]))
            .unwrap()
            .unwrap();
        after.set_cell_value_tracked(0, 0, 0, "=SEQUENCE(IF(B3=\"Area\",1,5))");
        assert!(prepare_header_replay(&after, &commit, true).is_err());
        assert_eq!(after.active_sheet().get_raw(2, 1), "Area");
    }

    #[test]
    fn rewind_replays_header_names_and_saved_projection() {
        let base = fixture(true);
        let (after, commit) = prepare_header_paste(&base, 0, 2, 1, &grid(&["Area", "Revenue"]))
            .unwrap()
            .unwrap();
        let mut history = History::new();
        history.record_action_with_provenance(
            UndoAction::TableCommit {
                sheet_index: 0,
                commit: Box::new(commit),
                description: "Paste headers".into(),
            },
            None,
        );
        let preview = history
            .build_workbook_before(1, Some(&base), 100, 10_000)
            .unwrap();
        assert_eq!(
            preview.workbook.active_sheet().tables(),
            after.active_sheet().tables()
        );
        assert_eq!(projection(&preview.workbook), projection(&base));
        let before = history
            .build_workbook_before(0, Some(&base), 100, 10_000)
            .unwrap();
        assert_eq!(before.workbook.active_sheet().get_raw(2, 1), "Group");
        assert!(preview.view_state.per_sheet[0].table_rows.is_some());
    }

    #[test]
    fn native_roundtrip_preserves_renamed_headers_rules_and_criteria() {
        let mut before = fixture(true);
        let id = before.active_sheet().tables()[0].id;
        before
            .set_calculated_column(id, 3, 3, "=[@Amount]*2", true)
            .unwrap();
        let (after, _) = prepare_header_paste(&before, 0, 2, 1, &grid(&["Area", "Revenue"]))
            .unwrap()
            .unwrap();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("headers.sheet");
        visigrid_io::native::save_workbook(&after, &path).unwrap();
        let loaded = visigrid_io::native::load_workbook(&path).unwrap();
        assert_eq!(
            loaded.active_sheet().tables(),
            after.active_sheet().tables()
        );
        assert_eq!(
            loaded.active_sheet().table_view_spec(),
            after.active_sheet().table_view_spec()
        );
        assert_eq!(loaded.active_sheet().get_raw(3, 3), "=[@[Revenue]]*2");
        assert_eq!(projection(&loaded), projection(&after));
    }

    #[test]
    fn history_gate_only_admits_column_renames_and_header_only_tables_work() {
        let mut before = fixture(false);
        let id = before.active_sheet().tables()[0].id;
        let rename = before.rename_table(id, "Orders").unwrap();
        assert!(!is_header_rename(&rename));
        before.set_table_view_spec(SheetId(7), None).unwrap();
        let mut range = before.table(id).unwrap().1.range;
        range.end_row = range.start_row;
        let resize = before.resize_table(id, range).unwrap();
        assert!(!is_header_rename(&resize));
        let (after, commit) = prepare_header_paste(&before, 0, 2, 1, &grid(&["Area", "Revenue"]))
            .unwrap()
            .unwrap();
        assert!(is_header_rename(&commit));
        assert_eq!(after.table(id).unwrap().1.range, range);
        assert_eq!(after.active_sheet().get_raw(3, 1), "West");
    }
}
