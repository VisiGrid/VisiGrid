//! Sparse, atomic footer authoring. Body bounds and saved criteria never move.
use super::{HeaderCell, TableCommit, Workbook};
use crate::cell::CellValue;
use crate::formula::structured::{StructuredReference, TableSection};
use crate::table::{DataTable, TableId, TableTotal, TableTotals};
use std::collections::BTreeSet;

fn value(table: &DataTable, offset: usize, total: &TableTotal) -> Result<CellValue, String> {
    if let Some(formula) = &total.formula {
        if !formula.starts_with('=') || crate::formula::parser::parse(formula).is_err() {
            return Err("Enter a valid formula starting with =.".into());
        }
        return Ok(CellValue::from_input(formula));
    }
    let code = match total.function.as_deref() {
        Some("average") => 101,
        Some("countNums") => 102,
        Some("count") => 103,
        Some("max") => 104,
        Some("min") => 105,
        Some("stdDev") => 107,
        Some("sum") => 109,
        Some("var") => 110,
        None | Some("none") => {
            return Ok(total
                .label
                .as_ref()
                .map_or(CellValue::Empty, |label| CellValue::Text(label.clone())))
        }
        Some("custom") => return Err("A custom total needs a formula.".into()),
        _ => return Err("Unknown totals function.".into()),
    };
    let column = table.columns[offset].name.clone();
    let reference = StructuredReference {
        table: None,
        section: TableSection::Data,
        columns: Some((column.clone(), column)),
    }
    .format();
    Ok(CellValue::from_input(&format!(
        "=SUBTOTAL({code},{reference})"
    )))
}

impl Workbook {
    /// Synchronize a host's canonical manual row visibility with every totals
    /// definition on the sheet. Criteria masks are separate and never changed.
    pub fn prepare_table_row_visibility(
        &self, sheet_id: crate::sheet::SheetId, hidden: BTreeSet<usize>,
    ) -> Result<(Workbook, super::super::GuardedStructureCommit), String> {
        self.ensure_writable()?;
        let sheet = self.sheet_by_id(sheet_id).ok_or("Visibility sheet no longer exists.")?;
        if hidden.iter().any(|row| *row >= crate::sheet::NUM_ROWS) {
            return Err("Hidden rows exceed the worksheet boundary.".into());
        }
        let mut candidate = self.clone();
        let legacy: BTreeSet<_> = hidden.iter().copied().filter(|row| *row < sheet.rows).collect();
        let changed = sheet.manual_hidden_rows != hidden || sheet.tables().iter()
            .any(|t| t.totals.as_ref().is_some_and(|t| t.hidden_rows != legacy));
        if changed {
            let values: Vec<_> = self.sheets().iter().flat_map(|sheet| {
                sheet.cells_iter().filter_map(move |((row, col), cell)| {
                    matches!(cell.value(), crate::cell::ValueRef::Formula { .. })
                        .then(|| (sheet.id, row, col, sheet.get_computed_value(row, col)))
                })
            }).collect();
            let sheet = candidate.sheet_by_id_mut(sheet_id).unwrap();
            sheet.set_manual_hidden_rows(hidden)?;
            if !super::super::recalc_deferred() {
                candidate.rebuild_dep_graph();
                candidate.recompute_full_ordered();
            }
            let changed: std::collections::HashSet<_> = values.into_iter().filter_map(|(id, row, col, before)| {
                (candidate.sheet_by_id(id)?.get_computed_value(row, col) != before).then_some(id)
            }).collect();
            for id in changed { candidate.sheet_by_id_mut(id).unwrap().mark_table_changed(); }
            candidate.bump_revision_for_structure();
        }
        if let Some(error) = candidate.take_incremental_errors().first() {
            return Err(format!("Could not recalculate row visibility: {error:?}"));
        }
        let mut commit = self.capture_guarded_batch(&candidate)?;
        commit.sheet = sheet_id;
        Ok((candidate, commit))
    }

    /// Show a footer only in empty cells. Hiding clears its values, retaining
    /// settings for the next show and exact values in the undo commit. Manual
    /// worksheet hiding is independent: creating a footer never unhides its row.
    /// `hidden_rows` is the host's manual visibility, never its filter mask.
    pub fn set_table_totals_visible(
        &mut self,
        id: TableId,
        visible: bool,
        hidden_rows: BTreeSet<usize>,
    ) -> Result<TableCommit, String> {
        self.ensure_writable()?;
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        let old = old.clone();
        if old.totals_row().is_some() == visible {
            return Err("The totals row already has that visibility.".into());
        }
        let mut new = old.clone();
        let mut totals = old.totals.clone().unwrap_or_else(|| {
            let mut columns = vec![TableTotal::default(); old.columns.len()];
            if columns.len() > 1 {
                columns[0].label = Some("Total".into());
            }
            columns.last_mut().unwrap().function = Some("sum".into());
            TableTotals {
                visible: false,
                shown: Some(false),
                hidden_rows: BTreeSet::new(),
                columns,
            }
        });
        if !visible {
            // Worksheet cells are authoritative on import. Capture departures
            // from the retained Excel settings before removing the footer.
            let sheet = self.sheet_by_id(sheet_id).unwrap();
            for (offset, total) in totals.columns.iter_mut().enumerate() {
                let cell = sheet
                    .get_cell(old.range.end_row + 1, old.range.start_col + offset)
                    .value;
                if value(&old, offset, total)
                    .is_ok_and(|expected| super::same_value(&expected, &cell))
                {
                    continue;
                }
                *total = match cell {
                    CellValue::Empty => TableTotal::default(),
                    CellValue::Text(label) => TableTotal { label: Some(label), ..Default::default() },
                    CellValue::Formula { source, .. } => TableTotal { function: Some("custom".into()), formula: Some(source), label: None },
                    _ => return Err("This totals row contains a literal value that cannot be retained as a total setting. Choose a function, label or custom formula before hiding it.".into()),
                };
            }
        }
        totals.visible = visible;
        totals.shown = Some(visible);
        totals.hidden_rows = hidden_rows;
        new.totals = Some(totals);
        self.commit_table_totals(sheet_id, old, new, None)
    }

    /// Set one footer column. Labels are literal text, including leading '='.
    pub fn set_table_total(
        &mut self,
        id: TableId,
        col: usize,
        total: TableTotal,
    ) -> Result<TableCommit, String> {
        self.ensure_writable()?;
        let (sheet_id, old) = self.table(id).ok_or("Table no longer exists.")?;
        let old = old.clone();
        if old.totals_row().is_none() {
            return Err("Show the totals row first.".into());
        }
        if col < old.range.start_col || col > old.range.end_col {
            return Err("Choose a Table column.".into());
        }
        let offset = col - old.range.start_col;
        value(&old, offset, &total)?;
        let mut new = old.clone();
        new.totals.as_mut().unwrap().columns[offset] = total;
        self.commit_table_totals(sheet_id, old, new, Some(offset))
    }

    fn commit_table_totals(
        &mut self,
        sheet_id: crate::sheet::SheetId,
        old: DataTable,
        new: DataTable,
        column: Option<usize>,
    ) -> Result<TableCommit, String> {
        let sheet = self.sheet_by_id(sheet_id).unwrap();
        new.validate(sheet.rows, sheet.cols)?;
        let row = old.range.end_row + 1;
        let region = crate::table::TableRange {
            start_row: row,
            end_row: row,
            start_col: old.range.start_col,
            end_col: old.range.end_col,
        };
        self.validate_table_region(sheet_id, region, Some(old.id))?;
        let showing = old.totals_row().is_none() && new.totals_row().is_some();
        if showing {
            self.validate_empty_table_append(sheet_id, region)
                .map_err(|_| "The row below this Table contains data or comments. Clear it or move the data before showing totals.".to_string())?;
        }
        let mut cells = Vec::new();
        for offset in 0..old.columns.len() {
            if column.is_some_and(|selected| selected != offset) {
                continue;
            }
            let col = old.range.start_col + offset;
            let before = HeaderCell {
                row,
                col,
                value: sheet.get_cell(row, col).value,
            };
            let after = HeaderCell {
                row,
                col,
                value: if new.totals_row().is_some() {
                    value(&new, offset, &new.totals.as_ref().unwrap().columns[offset])?
                } else {
                    CellValue::Empty
                },
            };
            cells.push((before, after));
        }
        let mut commit = self.table_commit(sheet_id, old.id, Some(old), Some(new))?;
        commit.totals_edit = true;
        // Hiding releases local footer references, but these cells are cleared
        // by this operation. Their exact sources belong to the cell delta,
        // not a second dependent-formula rewrite with different replay checks.
        commit.formulas.retain(|change| {
            !cells.iter().any(|(cell, _)| {
                change.cell.sheet == sheet_id
                    && change.cell.row == cell.row
                    && change.cell.col == cell.col
            })
        });
        commit.cells = cells;
        if showing {
            commit.append_region = Some(region);
        }
        self.capture_table_cell_absence(&mut commit);
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }
}
