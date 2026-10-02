//! Sparse, guarded Table cell history. No workbook or RowView is retained.
use visigrid_engine::{
    cell::{Cell, ValueRef},
    sheet::{Sheet, SheetId},
    table::DataTable,
    table_view::TableViewSpec,
    workbook::Workbook,
};

#[derive(Clone, Debug)]
pub struct TableCellPatch {
    pub row: usize,
    pub col: usize,
    pub before: Option<Cell>,
    pub after: Option<Cell>,
}

#[derive(Clone, Debug)]
pub struct TableCellsCommit {
    pub sheet: SheetId,
    tables: Vec<DataTable>,
    view: Option<TableViewSpec>,
    pub patches: Vec<TableCellPatch>,
}

fn image(sheet: &Sheet, row: usize, col: usize) -> Option<Cell> {
    sheet.get_cell_opt(row, col).map(|c| {
        let mut cell = c.to_cell();
        cell.clear_spill_state();
        cell
    })
}

fn same_cell(a: &Option<Cell>, b: &Option<Cell>) -> bool {
    match (a, b) {
        (None, None) => true,
        (Some(a), Some(b)) => {
            let value = match (a.value(), b.value()) {
                (ValueRef::Empty, ValueRef::Empty) => true,
                (ValueRef::Text(a), ValueRef::Text(b)) => a == b,
                (ValueRef::Number(a), ValueRef::Number(b)) => a.to_bits() == b.to_bits(),
                (ValueRef::Formula { source: a, .. }, ValueRef::Formula { source: b, .. }) => {
                    a == b
                }
                _ => false,
            };
            value
                && a.format == b.format
                && a.comment() == b.comment()
                && a.style_id() == b.style_id()
                && a.frozen_formula() == b.frozen_formula()
        }
        _ => false,
    }
}

impl TableCellsCommit {
    pub fn capture(
        before: &Sheet,
        after: &Sheet,
        targets: impl IntoIterator<Item = (usize, usize)>,
    ) -> Self {
        let view = before.table_view_spec().cloned();
        let tables = before.tables().to_vec();
        let patches = targets
            .into_iter()
            .collect::<std::collections::BTreeSet<_>>()
            .into_iter()
            .filter_map(|(row, col)| {
                let before = image(before, row, col);
                let after = image(after, row, col);
                (!same_cell(&before, &after)).then_some(TableCellPatch {
                    row,
                    col,
                    before,
                    after,
                })
            })
            .collect();
        Self {
            sheet: before.id,
            tables,
            view,
            patches,
        }
    }

    pub fn changes(&self) -> Vec<(usize, usize, String, String)> {
        self.patches
            .iter()
            .map(|p| {
                (
                    p.row,
                    p.col,
                    p.before
                        .as_ref()
                        .map(|c| c.value.raw_display())
                        .unwrap_or_default(),
                    p.after
                        .as_ref()
                        .map(|c| c.value.raw_display())
                        .unwrap_or_default(),
                )
            })
            .collect()
    }

    /// Hidden records must remain replayable: the original edit may itself
    /// have filtered them out. Validate canonical addresses, schema and exact
    /// expected cell images, then recalculate and validate all saved views.
    pub fn replay(&self, wb: &mut Workbook, undo: bool) -> Result<(), String> {
        wb.ensure_writable()?;
        let index = wb
            .sheet_index_by_id(self.sheet)
            .ok_or("The history sheet no longer exists.")?;
        let sheet = wb.sheet(index).unwrap();
        if sheet.table_view_spec() != self.view.as_ref()
            || sheet.tables().len() != self.tables.len()
            || self.tables.iter().any(|saved| {
                !sheet.tables().iter().any(|current| {
                    let mut saved = saved.clone();
                    let mut current = current.clone();
                    saved.next_column_id = 0;
                    current.next_column_id = 0;
                    saved == current
                })
            })
        {
            return Err(
                "The Table or its view changed since this edit. Undo/redo was not applied.".into(),
            );
        }
        crate::table_edit::validate_view_safe_targets(
            wb,
            index,
            &self
                .patches
                .iter()
                .map(|p| (p.row, p.col))
                .collect::<Vec<_>>(),
            true,
        )?;
        for patch in &self.patches {
            let expected = if undo { &patch.after } else { &patch.before };
            if !same_cell(&image(sheet, patch.row, patch.col), expected) {
                return Err(
                    "A target cell changed since this edit. Undo/redo was not applied.".into(),
                );
            }
        }
        if self.patches.is_empty() {
            return Ok(());
        }
        let mut candidate = wb.clone();
        {
            let mut batch = candidate.batch_guard();
            for patch in &self.patches {
                batch.restore_cell_tracked(
                    index,
                    patch.row,
                    patch.col,
                    if undo {
                        patch.before.clone()
                    } else {
                        patch.after.clone()
                    },
                )?;
            }
        }
        if let Some(error) = candidate.take_incremental_errors().first() {
            return Err(format!("Could not recalculate history: {error:?}"));
        }
        crate::table_edit::validate_view_safe_targets(
            &candidate,
            index,
            &self
                .patches
                .iter()
                .map(|p| (p.row, p.col))
                .collect::<Vec<_>>(),
            true,
        )?;
        for sheet in candidate.sheets() {
            sheet.build_saved_table_view(sheet.rows)?;
        }
        wb.restore_snapshot_monotonic(&candidate);
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::table_edit::{prepare_table_writes, tests::fixture, TableCellWrite};
    use visigrid_engine::{
        cell::{CellComment, CellValue},
        formula::eval::Value,
    };

    #[test]
    fn exact_cell_images_restore_absence_comments_style_and_frozen_formula() {
        let mut before = fixture(true);
        before.restore_cell_tracked(0, 3, 3, None).unwrap();
        let mut frozen = Cell::default();
        frozen.value = CellValue::Number(42.0);
        frozen.set_frozen_formula(Some("=REMOTE()".into()));
        frozen.set_style_id(Some(17));
        frozen.set_comment(Some(CellComment {
            text: "Keep this note".into(),
            author: "Tester".into(),
        }));
        before.restore_cell_tracked(0, 5, 3, Some(frozen)).unwrap();
        let mut write = TableCellWrite::value(5, 3, "=1+1".into());
        write.literal_text = true;
        write.comment = Some(None);
        let mut format = before.active_sheet().get_format(5, 3);
        format.bold = true;
        write.format = Some(format);
        let writes = vec![write, TableCellWrite::value(3, 3, "99".into())];
        let mut after = prepare_table_writes(&before, 0, &writes).unwrap();
        let commit = TableCellsCommit::capture(
            before.active_sheet(),
            after.active_sheet(),
            [(3, 3), (5, 3)],
        );
        assert_eq!(commit.patches.len(), 2);
        assert!(commit.patches[0].before.is_none());
        let original_after = image(after.active_sheet(), 5, 3);
        commit.replay(&mut after, true).unwrap();
        assert!(after.active_sheet().get_cell_opt(3, 3).is_none());
        assert!(same_cell(
            &image(after.active_sheet(), 5, 3),
            &image(before.active_sheet(), 5, 3)
        ));
        commit.replay(&mut after, false).unwrap();
        assert!(same_cell(
            &image(after.active_sheet(), 5, 3),
            &original_after
        ));
        assert_eq!(
            after.active_sheet().get_computed_value(5, 3),
            Value::Text("=1+1".into())
        );
        let mut changed_format = after.active_sheet().get_format(5, 3);
        changed_format.italic = true;
        after.sheet_mut(0).unwrap().set_format(5, 3, changed_format);
        let rev = after.revision();
        assert!(commit.replay(&mut after, true).is_err());
        assert_eq!(after.revision(), rev);
        assert_eq!(after.active_sheet().get_raw(3, 3), "99");
    }
}
