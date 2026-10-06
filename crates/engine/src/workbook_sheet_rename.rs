//! A sheet label and its authored references change in one guarded transaction.
use super::{GuardedStructureCommit, Workbook};
use crate::sheet::SheetId;

impl Workbook {
    pub fn prepare_sheet_rename(
        &self,
        id: SheetId,
        expected: &str,
        name: &str,
    ) -> Result<(Workbook, GuardedStructureCommit), String> {
        self.ensure_writable()?;
        let index = self
            .sheet_index_by_id(id)
            .ok_or("The sheet no longer exists.")?;
        if self.sheets[index].name != expected {
            return Err("The sheet name changed. Reopen rename and try again.".into());
        }
        let name = name.trim();
        if name.is_empty() {
            return Err("Sheet name cannot be empty.".into());
        }
        if name.chars().count() > 31 {
            return Err("Sheet name cannot exceed 31 characters.".into());
        }
        if !self.is_name_available(name, id) {
            return Err(format!("Sheet '{name}' already exists."));
        }
        if expected == name {
            return Ok((self.clone(), self.capture_guarded_batch(self)?));
        }
        let mut candidate = self.clone();
        candidate.rewrite_formula_sources(|source| {
            let result = crate::formula::sheets::rename_sheet_reference(source, expected, name)?;
            let changed = *source != result;
            *source = result;
            Ok(changed)
        })?;
        candidate.sheets[index].set_name(name);
        candidate.rebuild_dep_graph();
        let report = candidate.recompute_full_ordered();
        if (report.had_cycles && candidate.has_new_cycles(self))
            || report
                .errors
                .iter()
                .any(|e| e.error.contains("not settled"))
        {
            return Err(
                "Renaming would create a cycle or an unsettled calculation. Nothing was renamed."
                    .into(),
            );
        }
        for sheet in &mut candidate.sheets {
            sheet.mark_table_changed();
            sheet.build_saved_table_view(sheet.rows)?;
        }
        let mut commit = self.capture_sheet_rename(&candidate)?;
        commit.sheet = id;
        candidate.bump_revision_for_structure();
        Ok((candidate, commit))
    }
}
