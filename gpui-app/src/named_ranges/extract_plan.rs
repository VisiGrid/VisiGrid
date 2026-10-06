//! A captured extraction operates on visible canonical formula cells on one
//! sheet. The name and formula edits publish together through guarded history.
use crate::table_edit::{prepare_table_writes, validate_view_safe_targets, TableCellWrite};
use visigrid_engine::{
    cell::ValueRef,
    filter::RowView,
    formula::{
        extract::{first_reference_literal, reference_literal_count, replace_reference_literal},
        parser::{parse, Expr},
    },
    named_range::NamedRange,
    sheet::{SheetId, UnboundSheetRef},
    workbook::{GuardedStructureCommit, NamedRangeEdit, Workbook},
};

#[derive(Clone, Debug)]
pub(crate) struct ExtractionDraft {
    revision: u64,
    sheet: SheetId,
    generations: Vec<(SheetId, u64)>,
    formulas: Vec<(usize, usize, String)>,
    pub literal: String,
    pub occurrences: usize,
    range: NamedRange,
}

impl ExtractionDraft {
    pub fn capture(
        wb: &Workbook,
        rows: &RowView,
        selected: (usize, usize),
    ) -> Result<Self, String> {
        wb.ensure_writable()?;
        if !rows.is_view_row_visible(selected.0) {
            return Err("Select a visible formula cell.".into());
        }
        let sheet = wb.active_sheet();
        let row = rows.view_to_data(selected.0);
        let source = match sheet.get_cell_opt(row, selected.1).map(|c| c.value()) {
            Some(ValueRef::Formula { source, .. }) => source.to_string(),
            _ => return Err("Select a formula containing a cell or range reference.".into()),
        };
        let literal = first_reference_literal(&source)?
            .ok_or("No cell or bounded range reference was found in this formula.")?;
        let expr = parse(&format!("={literal}"))?;
        let index = |reference: UnboundSheetRef| -> Result<usize, String> {
            match reference {
                UnboundSheetRef::Current => Ok(wb.active_sheet_index()),
                UnboundSheetRef::Named(name) => wb
                    .sheet_id_by_name(&name)
                    .and_then(|id| wb.sheet_index_by_id(id))
                    .ok_or("The referenced sheet no longer exists.".into()),
            }
        };
        let range = match expr {
            Expr::CellRef {
                sheet, row, col, ..
            } => NamedRange::cell("", index(sheet)?, row, col),
            Expr::Range {
                sheet,
                start_row,
                start_col,
                end_row,
                end_col,
                ..
            } => NamedRange::range(
                "",
                index(sheet)?,
                start_row.min(end_row),
                start_col.min(end_col),
                start_row.max(end_row),
                start_col.max(end_col),
            ),
            _ => return Err("Choose a cell or bounded range reference.".into()),
        };
        let mut formulas = Vec::new();
        let mut occurrences = 0;
        for ((row, col), cell) in sheet.cells_iter() {
            if !rows.is_data_row_visible(row) {
                continue;
            }
            if let ValueRef::Formula { source, .. } = cell.value() {
                let count = reference_literal_count(source, &literal)?;
                if count == 0 {
                    continue;
                }
                if formulas.len() == 100_000 {
                    return Err("Extract from at most 100,000 visible formulas at a time.".into());
                }
                formulas.push((row, col, source.into()));
                occurrences += count;
            }
        }
        formulas.sort_by_key(|(r, c, _)| (*r, *c));
        validate_view_safe_targets(
            wb,
            wb.active_sheet_index(),
            &formulas
                .iter()
                .map(|(r, c, _)| (*r, *c))
                .collect::<Vec<_>>(),
            false,
        )?;
        Ok(Self {
            revision: wb.revision(),
            sheet: sheet.id,
            generations: wb
                .sheets()
                .iter()
                .map(|s| (s.id, s.edit_generation()))
                .collect(),
            formulas,
            literal,
            occurrences,
            range,
        })
    }

    pub fn cells(&self) -> Vec<(usize, usize)> {
        self.formulas.iter().map(|(r, c, _)| (*r, *c)).collect()
    }

    pub fn prepare(
        &self,
        wb: &Workbook,
        rows: &RowView,
        name: &str,
        description: Option<String>,
    ) -> Result<(Workbook, GuardedStructureCommit), String> {
        wb.ensure_writable()?;
        if wb.revision() != self.revision
            || wb.active_sheet_id() != self.sheet
            || wb
                .sheets()
                .iter()
                .map(|s| (s.id, s.edit_generation()))
                .collect::<Vec<_>>()
                != self.generations
        {
            return Err(
                "The workbook changed while this dialog was open. Reopen extraction and try again."
                    .into(),
            );
        }
        let index = wb.active_sheet_index();
        let mut writes = Vec::new();
        for (r, c, before) in &self.formulas {
            if !rows.is_data_row_visible(*r) || wb.active_sheet().get_raw(*r, *c) != *before {
                return Err("An affected formula changed or became hidden. Reopen extraction and try again.".into());
            }
            let (after, _) = replace_reference_literal(before, &self.literal, name)?;
            writes.push(TableCellWrite::value(*r, *c, after));
        }
        let mut range = self.range.clone();
        range.name = name.into();
        range.description = description;
        let (named, _) = wb.prepare_named_range_edit(&NamedRangeEdit::Create(range))?;
        let mut candidate = prepare_table_writes(&named, index, &writes)?;
        // Also explicit in manual-calculation mode; extraction must publish a
        // validated result and refresh every affected Table projection.
        candidate.rebuild_dep_graph();
        let report = candidate.recompute_full_ordered();
        if (report.had_cycles && candidate.has_new_cycles(wb))
            || report
                .errors
                .iter()
                .any(|e| e.error.contains("not settled"))
        {
            return Err(
                "Extraction would create a cycle or an unsettled calculation. Nothing was changed."
                    .into(),
            );
        }
        validate_view_safe_targets(&candidate, index, &self.cells(), true)?;
        let commit = wb.capture_guarded_batch(&candidate)?;
        Ok((candidate, commit))
    }
}

#[cfg(test)]
#[path = "extract_tests.rs"]
mod tests;
