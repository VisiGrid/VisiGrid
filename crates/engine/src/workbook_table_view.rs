//! Saved Table view intent and sparse, guarded history. No display mapping or
//! hidden rows are stored in the document. Hosts activate a rebuilt projection
//! only after their own range-owner, editing and layout checks pass.

use crate::{
    sheet::{Sheet, SheetId},
    table::DataTable,
    table_view::{validate_table_view_layout, TableView, TableViewSpec},
    workbook::Workbook,
};

/// History owns only before/after criteria. No cells, cached permutations or
/// visibility masks. Replay checks the current intent and target bindings.
#[derive(Clone, Debug)]
pub struct TableViewCommit {
    sheet: SheetId,
    before: Option<TableViewSpec>,
    after: Option<TableViewSpec>,
}

impl TableViewCommit {
    pub fn sheet_id(&self) -> SheetId {
        self.sheet
    }

    pub fn before(&self) -> Option<&TableViewSpec> {
        self.before.as_ref()
    }

    pub fn after(&self) -> Option<&TableViewSpec> {
        self.after.as_ref()
    }

    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }
}

impl Sheet {
    pub fn table_view_spec(&self) -> Option<&TableViewSpec> {
        self.table_view_spec.as_ref()
    }

    /// Validate only catalog bindings. A saved view can be suspended by an
    /// empty body or later neighboring content; neither discards its criteria.
    pub fn validate_table_view_spec(&self) -> Result<(), String> {
        self.validate_table_view_schema(self.tables())
    }

    pub(crate) fn validate_table_view_schema(&self, tables: &[DataTable]) -> Result<(), String> {
        if let Some(spec) = self.table_view_spec() {
            let table = tables
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("Clear the saved Table view before removing its Table.")?;
            spec.validate_schema(table).map_err(|e| {
                format!("Clear the affected Table view criterion before changing its schema: {e}")
            })?;
        }
        Ok(())
    }

    /// Rebuild from current computed values. Refuses unsafe layouts without
    /// mutating or dropping the saved criteria. Call after recalculation.
    pub fn build_saved_table_view(&self, row_count: usize) -> Result<Option<TableView>, String> {
        if let Some(spec) = &self.table_view_spec {
            let table = self
                .tables()
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("The saved Table view's Table no longer exists.")?;
            spec.validate_schema(table)?;
            // Button visibility and a cleared spec do not project worksheet rows.
            if !spec.has_criteria() {
                return Ok(None);
            }
            // Deleting the final record suspends projection, not saved intent.
            // Direct activation of a new empty-body view remains refused.
            if table.range.data_rows() == 0 {
                return Ok(None);
            }
        }

        self.table_view_spec
            .clone()
            .map(|spec| TableView::build(self, spec, row_count, None))
            .transpose()
    }
}

impl Workbook {
    /// Change one sheet's saved intent. Different owners must be cleared first;
    /// hosts also preflight any active worksheet-range owner before calling.
    /// Hiding buttons and clearing sort/filters use independent spec changes.
    pub fn set_table_view_spec(
        &mut self,
        sheet: SheetId,
        after: Option<TableViewSpec>,
    ) -> Result<TableViewCommit, String> {
        self.ensure_writable()?;
        let before = self
            .sheet_by_id(sheet)
            .ok_or("Table view sheet no longer exists.")?
            .table_view_spec
            .clone();
        if before
            .as_ref()
            .zip(after.as_ref())
            .is_some_and(|(a, b)| a.table != b.table)
        {
            return Err("Clear the current Table view before switching to another Table.".into());
        }
        let commit = TableViewCommit {
            sheet,
            before,
            after,
        };
        self.apply_table_view_commit(&commit, false)?;
        Ok(commit)
    }

    /// Apply or undo one criteria change atomically. A failed replay retains
    /// criteria, cells, formula caches, pivot freshness and workbook revision.
    pub fn apply_table_view_commit(
        &mut self,
        commit: &TableViewCommit,
        undo: bool,
    ) -> Result<(), String> {
        self.ensure_writable()?;
        let (expected, target) = if undo {
            (&commit.after, &commit.before)
        } else {
            (&commit.before, &commit.after)
        };
        let sheet = self
            .sheet_by_id(commit.sheet)
            .ok_or("Table view sheet no longer exists.")?;
        if &sheet.table_view_spec != expected {
            return Err("Table view changed since this operation was prepared.".into());
        }
        if let Some(spec) = target {
            let table = sheet
                .tables()
                .iter()
                .find(|t| t.id == spec.table)
                .ok_or("Table view's Table no longer exists.")?;
            spec.validate_schema(table)?;
            if spec.has_criteria() { validate_table_view_layout(sheet, spec.table)?; }
        }
        if !commit.is_noop() {
            self.sheet_by_id_mut(commit.sheet).unwrap().table_view_spec = target.clone();
            // Presentation changes dirty the document, without recomputation
            // or the sheet edit-generation bump that marks pivots stale.
            self.increment_revision();
        }
        Ok(())
    }

    /// Before exporting any view-bearing document, refuse malformed/dangling
    /// intent. Layout eligibility is checked when activating, not when saving.
    pub fn validate_table_view_specs(&self) -> Result<(), String> {
        for sheet in self.sheets() {
            sheet.validate_table_view_spec()?;
        }
        Ok(())
    }
}
