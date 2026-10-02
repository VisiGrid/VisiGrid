//! Desktop Table view controls. View changes are saved and undoable.
use crate::{app::Spreadsheet, history::UndoAction};
use gpui::*;
use std::collections::HashSet;
use visigrid_engine::{
    filter::{ColumnFilter, FilterState, RowView, SortDirection, UniqueValueEntry},
    table::{DataTable, TableColumnId, TableId},
    table_view::{TableFilter, TableSort, TableViewSpec},
    workbook::{TableViewCommit, Workbook},
};

pub(crate) const TABLE_VIEW_EDIT_MESSAGE: &str =
    "Clear Table sorting and filters before this operation. Safe cell edits and pastes, Table-body fills, and header-name paste are supported.";
const VALUE_LIMIT: usize = 500;

pub(crate) fn has_table_criteria(wb: &Workbook) -> bool {
    wb.sheets().iter().any(|s| {
        s.table_view_spec()
            .is_some_and(|v| v.sort.is_some() || !v.filters.is_empty())
    })
}

pub(crate) struct TableFilterDropdown {
    pub table: TableId,
    pub column: TableColumnId,
    pub name: String,
    pub table_name: String,
    pub values: Vec<UniqueValueEntry>,
    pub checked: HashSet<usize>,
    pub search: String,
    pub limited: bool,
    pub error: Option<String>,
    pub anchor: (f32, f32),
    pub revision: u64,
}

pub(crate) fn desktop_layout_error(
    table: &DataTable,
    heights: Option<&std::collections::HashMap<usize, f32>>,
    hidden: Option<&std::collections::BTreeSet<usize>>,
    frozen_rows: usize,
) -> Option<String> {
    let body = table.range.start_row + 1..=table.range.end_row;
    if heights.is_some_and(|h| h.keys().any(|r| body.contains(r)))
        || hidden.is_some_and(|h| h.iter().any(|r| body.contains(r)))
        || (frozen_rows > table.range.start_row + 1 && frozen_rows <= table.range.end_row)
    {
        Some("Table views need uniform, visible body rows with no freeze boundary through the records. Reset row heights, unhide rows or unfreeze the body first.".into())
    } else {
        None
    }
}

impl Spreadsheet {
    pub(crate) fn validate_saved_view_layout(&self, wb: &Workbook) -> Result<(), String> {
        for sheet in wb.sheets() {
            if let Some(spec) = sheet.table_view_spec() {
                if let Some(table) = sheet.tables().iter().find(|t| t.id == spec.table) {
                    if let Some(error) = crate::table_filter_ui::desktop_layout_error(
                        table,
                        self.row_heights.get(&sheet.id),
                        self.hidden_rows.get(&sheet.id),
                        sheet.frozen_panes.0,
                    ) {
                        return Err(error);
                    }
                }
            }
        }
        Ok(())
    }

    pub(crate) fn table_view_record_counts(&self, cx: &App) -> Option<(usize, usize)> {
        if !self.table_view_installed {
            return None;
        }
        let spec = self.sheet(cx).table_view_spec()?;
        let table = self
            .sheet(cx)
            .tables()
            .iter()
            .find(|t| t.id == spec.table)?;
        let total = table.range.data_rows();
        let hidden = self.row_view.row_count() - self.row_view.visible_count();
        Some((total.saturating_sub(hidden), total))
    }

    pub(crate) fn block_table_view_edit(&mut self, cx: &mut Context<Self>) -> bool {
        if has_table_criteria(self.wb(cx)) {
            self.status_message = Some(TABLE_VIEW_EDIT_MESSAGE.into());
            cx.notify();
            true
        } else {
            false
        }
    }

    pub(crate) fn table_header_button(&self, row: usize, col: usize, cx: &App) -> Option<TableId> {
        if self.is_previewing() || self.review_mode.is_some() || self.recovery_warning.is_some() {
            return None;
        }
        let table = self.sheet(cx).table_header_at(row, col)?;
        let show = self
            .sheet(cx)
            .table_view_spec()
            .filter(|v| v.table == table.id)
            .is_none_or(|v| v.show_filter_buttons);
        show.then_some(table.id)
    }

    pub(crate) fn table_layout_check(&self, table: &DataTable) -> Result<(), String> {
        desktop_layout_error(
            table,
            self.row_heights.get(&self.cached_sheet_id()),
            self.hidden_rows.get(&self.cached_sheet_id()),
            self.view_state.frozen_rows,
        )
        .map_or(Ok(()), Err)
    }

    /// Called on open, sheet switch and revision changes. No document mutation.
    pub(crate) fn sync_table_view(&mut self, cx: &mut Context<Self>) {
        if self.is_previewing() || self.review_mode.is_some() {
            return;
        }
        let sheet = self.sheet(cx);
        let key = (sheet.id, self.wb(cx).revision());
        if self.table_view_sync_key == Some(key) {
            return;
        }
        let same_sheet = self.table_view_sync_key.is_some_and(|k| k.0 == key.0);
        let record = Some(if same_sheet {
            self.row_view.view_to_data(self.view_state.selected.0)
        } else {
            self.view_state.selected.0
        });
        let split_record = self.split_pane.as_ref().map(|p| {
            if same_sheet {
                self.row_view.view_to_data(p.view_state.selected.0)
            } else {
                p.view_state.selected.0
            }
        });
        let has_spec = sheet.table_view_spec().is_some();
        if has_spec || self.table_view_installed {
            let result = (|| {
                if let Some(spec) = sheet.table_view_spec() {
                    if let Some(table) = sheet.tables().iter().find(|t| t.id == spec.table) {
                        self.table_layout_check(table)?;
                    }
                }
                sheet.build_saved_table_view(crate::app::NUM_ROWS)
            })();
            let row_count = crate::app::NUM_ROWS;
            match result {
                Ok(Some(view)) => {
                    if let Some(record) = record {
                        if let Ok(focus) = view.focus_record(record) {
                            self.view_state
                                .select_cell(focus.view_row, self.view_state.selected.1);
                            if focus.record_hidden {
                                self.status_message = Some("The selected record is filtered out; moved to the nearest visible row.".into());
                            }
                        }
                    }
                    if let (Some(pane), Some(record)) = (&mut self.split_pane, split_record) {
                        if let Ok(focus) = view.focus_record(record) {
                            pane.view_state
                                .select_cell(focus.view_row, pane.view_state.selected.1);
                        }
                    }
                    self.row_view = view.rows().clone();
                    self.filter_state = view.filters().clone();
                }
                other => {
                    self.row_view = RowView::new(row_count);
                    self.filter_state = FilterState::default();
                    if let Some(record) = record {
                        self.view_state
                            .select_cell(record, self.view_state.selected.1);
                    }
                    if let (Some(pane), Some(record)) = (&mut self.split_pane, split_record) {
                        pane.view_state
                            .select_cell(record, pane.view_state.selected.1);
                    }
                    if let Err(error) = other {
                        self.status_message = Some(format!(
                            "Table view suspended: {error} Clear view from its header menu."
                        ));
                    }
                }
            }
            self.table_view_installed = has_spec;
            self.ensure_visible(cx);
        }
        self.table_view_sync_key = Some(key);
    }

    pub(crate) fn open_table_filter(
        &mut self,
        id: TableId,
        col: usize,
        anchor: (f32, f32),
        cx: &mut Context<Self>,
    ) {
        if self.block_if_previewing_only(cx) {
            return;
        }
        if !self.mode.is_navigation() {
            self.commit_pending_edit(cx);
        }
        if !self.mode.is_navigation() {
            return;
        }
        self.sync_table_view(cx);
        let Some(table) = self.sheet(cx).tables().iter().find(|t| t.id == id) else {
            return;
        };
        let Some(column) = col
            .checked_sub(table.range.start_col)
            .and_then(|c| table.columns.get(c))
        else {
            return;
        };
        let spec = self.sheet(cx).table_view_spec().filter(|v| v.table == id);
        let criteria = spec.and_then(|v| v.filters.iter().find(|f| f.column == column.id));
        let mut state = FilterState {
            filter_range: Some((
                table.range.start_row,
                table.range.start_col,
                table.range.end_row,
                table.range.end_col,
            )),
            ..Default::default()
        };
        let mut values = state
            .build_unique_values(
                col,
                |r, c| self.sheet(cx).get_computed_value(r, c),
                VALUE_LIMIT + 1,
            )
            .to_vec();
        values.sort_by(|a, b| {
            b.count
                .cmp(&a.count)
                .then_with(|| a.display.to_lowercase().cmp(&b.display.to_lowercase()))
                .then_with(|| format!("{:?}", a.key).cmp(&format!("{:?}", b.key)))
        });
        let limited = values.len() > VALUE_LIMIT;
        let checked = values
            .iter()
            .enumerate()
            .filter(|(_, entry)| {
                criteria.is_none_or(|f| {
                    f.criteria
                        .selected
                        .as_ref()
                        .is_none_or(|set| set.contains(&entry.key))
                })
            })
            .map(|(i, _)| i)
            .collect();
        self.table_filter_dropdown = Some(TableFilterDropdown {
            table: id,
            column: column.id,
            name: column.name.clone(),
            table_name: table.name.clone(),
            values,
            checked,
            search: String::new(),
            limited,
            anchor,
            revision: self.wb(cx).revision(),
            error: (table.range.data_rows() == 0)
                .then(|| "Add records to this Table before sorting or filtering.".into()),
        });
        self.filter_dropdown_col = None;
        cx.notify();
    }

    pub(crate) fn change_table_view(
        &mut self,
        spec: Option<TableViewSpec>,
        description: &str,
        cx: &mut Context<Self>,
    ) -> bool {
        if self.block_if_previewing_only(cx) {
            return false;
        }
        if self.import_in_progress || self.hub_activity.is_some() || !self.mode.is_navigation() {
            self.table_filter_error(
                "Finish the current edit or import before changing a Table view.".into(),
                cx,
            );
            return false;
        }
        if let Some(spec) = &spec {
            if !self.table_view_installed && self.filter_state.is_enabled() {
                self.table_filter_error(
                    "Clear the worksheet sort/filter before activating a Table view.".into(),
                    cx,
                );
                return false;
            }
            if let Some(table) = self.sheet(cx).tables().iter().find(|t| t.id == spec.table) {
                if let Err(error) = self.table_layout_check(table) {
                    self.table_filter_error(error, cx);
                    return false;
                }
            }
        }
        let sheet = self.sheet(cx).id;
        let result = self
            .workbook
            .update(cx, |wb, _| wb.set_table_view_spec(sheet, spec));
        match result {
            Err(error) => {
                self.table_filter_error(error, cx);
                false
            }
            Ok(commit) => {
                if !commit.is_noop() {
                    self.history.record_action_with_provenance(
                        UndoAction::TableViewChanged {
                            sheet_index: self.sheet_index(cx),
                            commit: Box::new(commit),
                            description: description.into(),
                        },
                        None,
                    );
                    self.is_modified = true;
                }
                self.table_filter_dropdown = None;
                self.sync_table_view(cx);
                self.status_message = Some(description.into());
                cx.notify();
                true
            }
        }
    }

    fn table_filter_error(&mut self, error: String, cx: &mut Context<Self>) {
        if let Some(menu) = &mut self.table_filter_dropdown {
            menu.error = Some(error.clone());
        }
        self.status_message = Some(error);
        cx.notify();
    }

    pub(crate) fn table_menu_action(&mut self, action: &str, cx: &mut Context<Self>) {
        let Some(menu) = &self.table_filter_dropdown else {
            return;
        };
        if menu.revision != self.wb(cx).revision() {
            self.table_filter_error(
                "The workbook changed. Close and reopen this menu.".into(),
                cx,
            );
            return;
        }
        let id = menu.table;
        if matches!(action, "ascending" | "descending" | "apply")
            && self
                .sheet(cx)
                .tables()
                .iter()
                .find(|t| t.id == id)
                .is_none_or(|t| t.range.data_rows() == 0)
        {
            self.table_filter_error(
                "Add records to this Table before sorting or filtering.".into(),
                cx,
            );
            return;
        }
        let mut spec = self
            .sheet(cx)
            .table_view_spec()
            .filter(|v| v.table == id)
            .cloned()
            .unwrap_or_else(|| TableViewSpec::new(id));
        let description = match action {
            "ascending" | "descending" => {
                spec.sort = Some(TableSort {
                    column: menu.column,
                    direction: if action == "ascending" {
                        SortDirection::Ascending
                    } else {
                        SortDirection::Descending
                    },
                });
                format!(
                    "Sort {} by {} {}",
                    menu.table_name,
                    menu.name,
                    if action == "ascending" {
                        "ascending"
                    } else {
                        "descending"
                    }
                )
            }
            "clear-sort" => {
                spec.clear_sort();
                "Clear Table sort".into()
            }
            "clear-filters" => {
                spec.clear_filters();
                "Clear Table filters".into()
            }
            "clear-view" => {
                if self
                    .sheet(cx)
                    .table_view_spec()
                    .is_some_and(|v| v.table != id)
                {
                    self.table_filter_error(
                        "Open the active Table's header menu to clear its view.".into(),
                        cx,
                    );
                    return;
                }
                self.change_table_view(None, "Clear Table view — editing enabled", cx);
                return;
            }
            "apply" => {
                if menu.limited {
                    self.table_filter_error("More than 500 unique values. Value filtering is unavailable for this column; sorting is still available.".into(), cx);
                    return;
                }
                spec.filters.retain(|f| f.column != menu.column);
                if menu.checked.len() != menu.values.len() {
                    spec.filters.push(TableFilter {
                        column: menu.column,
                        criteria: ColumnFilter {
                            selected: Some(
                                menu.checked
                                    .iter()
                                    .filter_map(|i| menu.values.get(*i))
                                    .map(|e| e.key.clone())
                                    .collect(),
                            ),
                            text_filter: None,
                        },
                    });
                }
                format!("Filter {} · {}", menu.table_name, menu.name)
            }
            _ => return,
        };
        self.change_table_view(Some(spec), &description, cx);
    }

    pub(crate) fn replay_table_view(
        &mut self,
        commit: &TableViewCommit,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        if let Some(spec) = if undo {
            commit.before()
        } else {
            commit.after()
        } {
            if let Some(table) = self
                .wb(cx)
                .sheet_by_id(commit.sheet_id())
                .and_then(|s| s.tables().iter().find(|t| t.id == spec.table))
            {
                if let Some(error) = desktop_layout_error(
                    table,
                    self.row_heights.get(&commit.sheet_id()),
                    self.hidden_rows.get(&commit.sheet_id()),
                    self.wb(cx)
                        .sheet_by_id(commit.sheet_id())
                        .map_or(0, |s| s.frozen_panes.0),
                ) {
                    self.status_message = Some(error);
                    cx.notify();
                    return false;
                }
            }
        }
        let result = self
            .workbook
            .update(cx, |wb, _| wb.apply_table_view_commit(commit, undo));
        if let Err(error) = result {
            self.status_message = Some(error);
            cx.notify();
            return false;
        }
        if let Some(index) = self.wb(cx).sheet_index_by_id(commit.sheet_id()) {
            if index != self.sheet_index(cx) {
                self.activate_sheet(index, cx);
            }
        }
        self.table_filter_dropdown = None;
        self.sync_table_view(cx);
        cx.notify();
        true
    }

    pub(crate) fn table_filter_key(&mut self, key: &Keystroke, cx: &mut Context<Self>) -> bool {
        let Some(menu) = &mut self.table_filter_dropdown else {
            return false;
        };
        match key.key.as_str() {
            "escape" => self.table_filter_dropdown = None,
            "enter" => {
                self.table_menu_action("apply", cx);
                return true;
            }
            "backspace" => {
                menu.search.pop();
            }
            _ => {
                if !key.modifiers.control && !key.modifiers.platform && !key.modifiers.alt {
                    if let Some(text) = &key.key_char {
                        menu.search.push_str(text);
                    }
                }
            }
        }
        cx.notify();
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use visigrid_engine::{
        sheet::{Sheet, SheetId},
        table::TableRange,
    };

    fn fixture() -> (Workbook, TableViewSpec) {
        let mut wb = Workbook::from_sheets(vec![Sheet::new(SheetId(7), 30, 8)], 0);
        for (r, value) in ["Amount", "20", "10", "30"].iter().enumerate() {
            wb.set_cell_value_tracked(0, r + 2, 1, value);
        }
        let id = wb
            .create_table(
                SheetId(7),
                TableRange {
                    start_row: 2,
                    start_col: 1,
                    end_row: 5,
                    end_col: 1,
                },
                "Sales",
            )
            .unwrap()
            .table_id();
        (wb, TableViewSpec::new(id))
    }

    #[::core::prelude::v1::test]
    fn table_view_edit_gate_covers_other_sheets_and_hidden_buttons() {
        let (mut wb, mut spec) = fixture();
        wb.set_table_view_spec(SheetId(7), Some(spec.clone()))
            .unwrap();
        assert!(
            !has_table_criteria(&wb),
            "buttons alone must not block editing"
        );
        spec.sort = Some(TableSort {
            column: wb.table(spec.table).unwrap().1.columns[0].id,
            direction: SortDirection::Ascending,
        });
        spec.show_filter_buttons = false;
        wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
        wb.add_sheet();
        wb.set_active_sheet(1);
        assert!(has_table_criteria(&wb));
        wb.set_table_view_spec(SheetId(7), None).unwrap();
        assert!(!has_table_criteria(&wb));
    }

    #[::core::prelude::v1::test]
    fn table_view_layout_checks_body_metadata_and_freeze_boundary() {
        let (wb, spec) = fixture();
        let table = wb.table(spec.table).unwrap().1;
        assert!(
            desktop_layout_error(table, Some(&[(2, 32.0)].into()), Some(&[1].into()), 3).is_none()
        );
        assert!(desktop_layout_error(table, Some(&[(3, 32.0)].into()), None, 0).is_some());
        assert!(desktop_layout_error(table, None, Some(&[5].into()), 0).is_some());
        assert!(desktop_layout_error(table, None, None, 4).is_some());
        assert!(desktop_layout_error(table, None, None, 6).is_none());
    }

    #[::core::prelude::v1::test]
    fn table_view_lua_batch_rejects_before_any_cell_changes() {
        use visigrid_scripting::{LuaCellValue, LuaOp};
        let (mut wb, mut spec) = fixture();
        spec.sort = Some(TableSort {
            column: wb.table(spec.table).unwrap().1.columns[0].id,
            direction: SortDirection::Ascending,
        });
        wb.set_table_view_spec(SheetId(7), Some(spec)).unwrap();
        let revision = wb.revision();
        let ops = [
            LuaOp::SetValue {
                row: 0,
                col: 0,
                value: LuaCellValue::Number(99.0),
            },
            LuaOp::SetValue {
                row: 2,
                col: 1,
                value: LuaCellValue::String("bad header".into()),
            },
            LuaOp::ClearCell { row: 3, col: 1 },
        ];
        let error =
            crate::views::lua_console::apply_captured_lua_ops(&mut wb, 0, &ops).unwrap_err();
        assert_eq!(error, TABLE_VIEW_EDIT_MESSAGE);
        assert_eq!(wb.revision(), revision);
        assert_eq!(wb.active_sheet().get_raw(0, 0), "");
        assert_eq!(wb.active_sheet().get_raw(2, 1), "Amount");
        assert_eq!(wb.active_sheet().get_raw(3, 1), "20");
    }

    #[::core::prelude::v1::test]
    fn table_view_history_supports_undo_and_rewind() {
        let (mut wb, mut spec) = fixture();
        spec.sort = Some(TableSort {
            column: wb.table(spec.table).unwrap().1.columns[0].id,
            direction: SortDirection::Ascending,
        });
        let commit = wb
            .set_table_view_spec(SheetId(7), Some(spec.clone()))
            .unwrap();
        let mut history = crate::history::History::new();
        history.record_action_with_provenance(
            UndoAction::TableViewChanged {
                sheet_index: 0,
                commit: Box::new(commit),
                description: "Sort Table".into(),
            },
            None,
        );
        let entry = history.undo().unwrap();
        assert!(entry.action.is_replay_supported());
        let UndoAction::TableViewChanged { commit, .. } = entry.action else {
            panic!("wrong history action")
        };
        wb.apply_table_view_commit(&commit, true).unwrap();
        assert!(!has_table_criteria(&wb));
        history.redo().unwrap();
        wb.apply_table_view_commit(&commit, false).unwrap();
        assert_eq!(wb.active_sheet().table_view_spec(), Some(&spec));
        assert_eq!(wb.active_sheet().get_raw(3, 1), "20");
    }
}
