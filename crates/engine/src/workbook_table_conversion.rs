//! Extra authored metadata checks for Convert to Range. Called on a candidate.
use super::Workbook;
use crate::{formula::structured, sheet::SheetId, table::DataTable, validation::CellRange};

impl Workbook {
    pub(super) fn prepare_conversion_metadata(
        &mut self,
        owner: SheetId,
        table: &DataTable,
    ) -> Result<(), String> {
        let full = table.full_range();
        let references = |sheet: SheetId, range: &CellRange, source: &str| {
            structured::source_references(source)
                .iter()
                .any(|(_, _, r)| match &r.table {
                    Some(name) => name.eq_ignore_ascii_case(&table.name),
                    None => {
                        sheet == owner
                            && range.start_row <= full.end_row
                            && range.end_row >= full.start_row
                            && range.start_col <= full.end_col
                            && range.end_col >= full.start_col
                    }
                })
        };
        let mut frozen = Vec::new();
        for sheet in self.sheets() {
            // Rule predicates are evaluated at every target cell. Rewriting
            // only their first anchor would freeze [@Column] to one record.
            // Keep this refusal explicit until range-wise conversion exists.
            if sheet.cond_formats.iter().any(|r| {
                r.ranges
                    .iter()
                    .any(|range| references(sheet.id, range, &r.predicate))
            }) {
                return Err("A conditional-format rule refers to this Table. Change that rule to cell references before converting the Table to a range.".into());
            }
            for (range, rule) in sheet.validations.iter() {
                let mut referenced = false;
                rule.clone().map_references(|source| {
                    referenced |= references(sheet.id, range, source);
                    source.to_owned()
                });
                if referenced {
                    return Err("A validation rule refers to this Table. Change that rule to cell references before converting the Table to a range.".into());
                }
            }
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
