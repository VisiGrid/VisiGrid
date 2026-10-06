//! Sheet creation/deletion prepared on a private, recalculated candidate.
use super::{GuardedStructureCommit, Workbook};
use crate::sheet::SheetId;

impl Workbook {
    pub fn prepare_sheet_add(
        &self,
        name: Option<&str>,
    ) -> Result<(Workbook, GuardedStructureCommit), String> {
        self.ensure_writable()?;
        let mut candidate = self.clone();
        let index = if let Some(name) = name {
            let name = name.trim();
            if name.is_empty() || name.chars().count() > 31 {
                return Err("Choose a sheet name between 1 and 31 characters.".into());
            }
            candidate
                .add_sheet_named(name)
                .ok_or("A sheet with this name already exists.")?
        } else {
            candidate.add_sheet()
        };
        candidate.validate_sheet_lifecycle(self)?;
        let mut commit = self.capture_sheet_change(&candidate, index, true)?;
        commit.sheet = candidate.sheets[index].id;
        candidate.bump_revision_for_structure();
        Ok((candidate, commit))
    }

    pub fn prepare_sheet_delete(
        &self,
        id: SheetId,
    ) -> Result<(Workbook, GuardedStructureCommit), String> {
        self.ensure_writable()?;
        let index = self
            .sheet_index_by_id(id)
            .ok_or("The sheet no longer exists.")?;
        if self.sheets.len() == 1 {
            return Err("Cannot delete the last sheet.".into());
        }
        if self
            .sheets
            .iter()
            .filter(|s| s.id != id)
            .any(|s| s.pivots.iter().any(|p| p.source.sheet_id == id))
        {
            return Err("This sheet is a source for a PivotTable on another sheet. Remove the dependent PivotTable before deleting its source.".into());
        }
        let name = &self.sheets[index].name;
        let tables: Vec<_> = self.sheets[index]
            .tables()
            .iter()
            .map(|t| t.name.clone())
            .collect();
        let mut candidate = self.clone();
        let active = candidate.active_sheet_id();
        candidate.sheets.remove(index);
        candidate.active_sheet = candidate
            .sheet_index_by_id(active)
            .unwrap_or(index.min(candidate.sheets.len() - 1));
        candidate.named_ranges.remove_sheet(index);
        candidate.refresh_table_name_reservations();
        candidate.rewrite_formula_sources(|source| {
            let result = crate::formula::sheets::delete_sheet_references(source, name, &tables)?;
            let changed = *source != result;
            *source = result;
            Ok(changed)
        })?;
        candidate.validate_sheet_lifecycle(self)?;
        let mut commit = self.capture_sheet_change(&candidate, index, false)?;
        commit.sheet = id;
        candidate.bump_revision_for_structure();
        Ok((candidate, commit))
    }

    fn validate_sheet_lifecycle(&mut self, before: &Workbook) -> Result<(), String> {
        self.rebuild_dep_graph();
        let report = self.recompute_full_ordered();
        if (report.had_cycles && self.has_new_cycles(before))
            || report
                .errors
                .iter()
                .any(|e| e.error.contains("not settled"))
        {
            return Err("The sheet change would create a cycle or an unsettled calculation. Nothing was changed.".into());
        }
        for sheet in &mut self.sheets {
            sheet.mark_table_changed();
            sheet.build_saved_table_view(sheet.rows)?;
        }
        Ok(())
    }
}
