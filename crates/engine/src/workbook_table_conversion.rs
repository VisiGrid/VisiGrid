//! Extra authored metadata checks for Convert to Range. Called on a candidate.
use super::Workbook;
use crate::{sheet::SheetId, table::DataTable};

#[path = "workbook_table_conversion_rules.rs"]
mod rules;

impl Workbook {
    pub(super) fn prepare_conversion_metadata(
        &mut self,
        owner: SheetId,
        table: &DataTable,
    ) -> Result<(), String> {
        self.rewrite_conversion_rules(owner, table)?;
        let mut frozen = Vec::new();
        for sheet in self.sheets() {
            for ((row, col), cell) in sheet.cells_iter() {
                let Some(source) = cell.frozen_formula() else {
                    continue;
                };
                let rewritten = self
                    .rewrite_table_formula_source(owner, table, None, sheet.id, row, col, source)?;
                if rewritten != source {
                    let mut cell = cell.to_cell();
                    cell.set_frozen_formula(Some(rewritten));
                    frozen.push((sheet.id, row, col, cell));
                }
            }
        }
        for (sheet, row, col, cell) in frozen {
            self.sheet_by_id_mut(sheet)
                .unwrap()
                .restore_history_cell(row, col, Some(cell));
        }
        Ok(())
    }
}
