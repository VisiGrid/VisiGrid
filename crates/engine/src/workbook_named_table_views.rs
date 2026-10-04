//! Named criteria presets share the Table's identity, validation and sparse
//! history. Updating a preset never changes the active sheet view.
use super::{TableCommit, Workbook};
use crate::{
    table::TableId,
    table_view::{NamedTableView, TableViewSpec},
};

impl Workbook {
    fn change_named_table_views(
        &mut self,
        id: TableId,
        edit: impl FnOnce(&mut Vec<NamedTableView>) -> Result<(), String>,
    ) -> Result<TableCommit, String> {
        self.ensure_writable()?;
        let (sheet, old) = self.table(id).ok_or("Table no longer exists.")?;
        let mut new = old.clone();
        edit(&mut new.saved_views)?;
        new.validate(crate::sheet::NUM_ROWS, crate::sheet::NUM_COLS)?;
        let commit = self.table_commit(sheet, id, Some(old.clone()), Some(new))?;
        self.apply_table_commit(&commit, false)?;
        Ok(commit)
    }

    /// Duplicate names are rejected. Replacing an existing preset is a
    /// separate, explicit operation.
    pub fn save_named_table_view(
        &mut self,
        id: TableId,
        name: &str,
        view: TableViewSpec,
    ) -> Result<TableCommit, String> {
        let name = name.trim().to_owned();
        self.change_named_table_views(id, |views| {
            if views.iter().any(|v| v.name.to_lowercase() == name.to_lowercase()) {
                return Err("A view with that name already exists. Choose another name or update the existing view.".into());
            }
            views.push(NamedTableView { name, view });
            Ok(())
        })
    }

    pub fn update_named_table_view(
        &mut self,
        id: TableId,
        name: &str,
        view: TableViewSpec,
    ) -> Result<TableCommit, String> {
        self.change_named_table_views(id, |views| {
            views
                .iter_mut()
                .find(|v| v.name == name)
                .ok_or("Saved view no longer exists.")?
                .view = view;
            Ok(())
        })
    }

    pub fn rename_named_table_view(
        &mut self,
        id: TableId,
        name: &str,
        new_name: &str,
    ) -> Result<TableCommit, String> {
        self.change_named_table_views(id, |views| {
            views
                .iter_mut()
                .find(|v| v.name == name)
                .ok_or("Saved view no longer exists.")?
                .name = new_name.trim().into();
            Ok(())
        })
    }

    pub fn delete_named_table_view(
        &mut self,
        id: TableId,
        name: &str,
    ) -> Result<TableCommit, String> {
        self.change_named_table_views(id, |views| {
            let index = views
                .iter()
                .position(|v| v.name == name)
                .ok_or("Saved view no longer exists.")?;
            views.remove(index);
            Ok(())
        })
    }
}
