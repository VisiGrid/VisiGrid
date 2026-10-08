//! Extra authored metadata checks for Convert to Range. Called on a candidate.
use super::Workbook;
use crate::{cell::CellValue, sheet::SheetId, table::DataTable};

#[path = "workbook_table_conversion_rules.rs"]
mod rules;

impl Workbook {
    pub(super) fn prepare_conversion_metadata(
        &mut self,
        owner: SheetId,
        table: &DataTable,
    ) -> Result<(), String> {
        self.rewrite_conversion_rules(owner, table)?;
        self.rewrite_conversion_totals(owner, table)?;
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

    fn rewrite_conversion_totals(
        &mut self,
        owner: SheetId,
        converted: &DataTable,
    ) -> Result<(), String> {
        let mut settings = Vec::new();
        let mut cells = Vec::new();
        for (sheet_id, table) in self.tables().filter(|(_, table)| table.id != converted.id) {
            let Some(totals) = &table.totals else {
                continue;
            };
            let mut rewritten = totals.clone();
            let row = table.range.end_row + 1;
            let sheet = self.sheet_by_id(sheet_id).unwrap();
            let mut changed_cells = false;
            for (offset, total) in rewritten.columns.iter_mut().enumerate() {
                let col = table.range.start_col + offset;
                if let Some(source) = &mut total.formula {
                    *source = self.rewrite_named_table_formula_source(
                        owner, converted, None, sheet_id, row, col, source,
                    )?;
                }
                if totals.visible {
                    let mut cell = sheet.get_cell(row, col);
                    if let CellValue::Formula { source, .. } = &cell.value {
                        let rewritten = self.rewrite_named_table_formula_source(
                            owner, converted, None, sheet_id, row, col, source,
                        )?;
                        if rewritten != *source {
                            cell.value = CellValue::from_input(&rewritten);
                            cells.push((sheet_id, row, col, cell));
                            changed_cells = true;
                        }
                    }
                }
            }
            if rewritten != *totals || changed_cells {
                self.validate_table_region(sheet_id, table.full_range(), Some(table.id))?;
                settings.push((sheet_id, table.id, rewritten));
            }
        }
        for (sheet_id, id, totals) in settings {
            let sheet = self.sheet_by_id_mut(sheet_id).unwrap();
            sheet
                .data_tables
                .iter_mut()
                .find(|table| table.id == id)
                .unwrap()
                .totals = Some(totals);
            sheet.mark_table_changed();
        }
        for (sheet_id, row, col, image) in cells {
            self.sheet_by_id_mut(sheet_id)
                .unwrap()
                .restore_history_cell(row, col, Some(image));
        }
        Ok(())
    }
}
