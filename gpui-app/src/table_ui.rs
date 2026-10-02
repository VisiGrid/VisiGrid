//! Desktop Table commands. Schema edits always use the engine's guarded,
//! sparse commits, including direct header edits and history replay.
use crate::{app::Spreadsheet, history::UndoAction};
use gpui::*;
use visigrid_engine::{
    sheet::SheetId,
    table::{DataTable, TableId, TableRange},
    workbook::TableCommit,
};

pub(crate) const TABLE_CONTROLS_HEIGHT: f32 = 32.0;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TableDialogKind {
    Create,
    Rename(TableId),
    Resize(TableId),
    Convert(TableId),
    ColumnFormula(TableId, usize, bool),
}

#[derive(Clone, Debug)]
pub(crate) struct TableDialog {
    pub kind: TableDialogKind,
    pub sheet: SheetId,
    pub name: String,
    pub range: String,
    pub has_headers: bool,
    pub field: usize,
    pub select_all: bool,
    pub error: Option<String>,
}

pub(crate) fn range_label(r: TableRange) -> String {
    format!(
        "{}{}:{}{}",
        Spreadsheet::col_letter(r.start_col),
        r.start_row + 1,
        Spreadsheet::col_letter(r.end_col),
        r.end_row + 1
    )
}

/// Only a finite, unqualified A1 rectangle; never quietly accept a name,
/// reversed range, whole column or another sheet.
pub(crate) fn parse_range(text: &str) -> Result<TableRange, String> {
    use visigrid_engine::formula::parser::{parse, Expr};
    let parsed = parse(&format!("={}", text.trim()))
        .map_err(|_| "Enter a range such as A1:D20.".to_string())?;
    let range = match parsed {
        Expr::CellRef {
            sheet: visigrid_engine::sheet::UnboundSheetRef::Current,
            row,
            col,
            ..
        } => TableRange {
            start_row: row,
            end_row: row,
            start_col: col,
            end_col: col,
        },
        Expr::Range {
            sheet: visigrid_engine::sheet::UnboundSheetRef::Current,
            start_row,
            start_col,
            end_row,
            end_col,
            ..
        } => TableRange {
            start_row,
            start_col,
            end_row,
            end_col,
        },
        _ => return Err("Enter a range on this sheet, such as A1:D20.".into()),
    };
    range.validate(crate::app::NUM_ROWS, crate::app::NUM_COLS)?;
    Ok(range)
}

fn header_in_view_rect(
    table: &DataTable,
    rows: &visigrid_engine::filter::RowView,
    r0: usize,
    c0: usize,
    r1: usize,
    c1: usize,
) -> bool {
    c0.min(c1) <= table.range.end_col
        && c0.max(c1) >= table.range.start_col
        && if !rows.is_sorted() {
            (r0.min(r1)..=r0.max(r1)).contains(&table.range.start_row)
        } else {
            (r0.min(r1)..=r0.max(r1)).any(|r| rows.view_to_data(r) == table.range.start_row)
        }
}

impl Spreadsheet {
    pub(crate) fn commit_calculated_value(
        &mut self,
        view_row: usize,
        col: usize,
        source: &str,
        cx: &mut Context<Self>,
    ) -> Option<bool> {
        let row = self.row_view.view_to_data(view_row);
        let sheet = self.sheet(cx).id;
        let result = self.workbook.update(cx, |wb, _| {
            wb.try_calculated_column(sheet, row, col, source)
        });
        match result {
            Ok(None) => None,
            Ok(Some(commit)) => {
                self.record_table_commit(commit, "Fill calculated column".into(), cx);
                Some(true)
            }
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
                Some(false)
            }
        }
    }

    pub(crate) fn restore_column_formula(&mut self, id: TableId, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) || self.mode.is_editing() {
            return;
        }
        let (row, col) = self.view_state.selected;
        let row = self.row_view.view_to_data(row);
        match self
            .workbook
            .update(cx, |wb, _| wb.restore_calculated_cell(id, row, col))
        {
            Ok(commit) => self.record_table_commit(commit, "Restore column formula".into(), cx),
            Err(error) => {
                self.status_message = Some(error);
                cx.notify();
            }
        }
    }

    fn table_growth_blocked(&mut self, cx: &mut Context<Self>) -> bool {
        if self.block_if_previewing(cx) {
            return true;
        }
        if self.row_view.is_sorted() || self.row_view.is_filtered() {
            self.status_message =
                Some("Clear sorting and filters before appending Table rows.".into());
            cx.notify();
            return true;
        }
        false
    }

    /// None means an ordinary edit; Some is an atomic append attempt.
    pub(crate) fn commit_table_append_value(
        &mut self,
        view_row: usize,
        col: usize,
        value: &str,
        cx: &mut Context<Self>,
    ) -> Option<bool> {
        if value.trim().is_empty() {
            return None;
        }
        let row = self.row_view.view_to_data(view_row);
        let range = TableRange {
            start_row: row,
            end_row: row,
            start_col: col,
            end_col: col,
        };
        let id = match self.wb(cx).table_append_target(self.sheet(cx).id, range) {
            Ok(Some(id)) => id,
            Ok(None) => return None,
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
                return Some(false);
            }
        };
        if self.table_growth_blocked(cx) {
            return Some(false);
        }
        let result = self.workbook.update(cx, |wb, _| {
            wb.append_table_rows_with_edit(id, 1, row, col, value)
        });
        Some(match result {
            Ok(commit) => {
                self.record_table_commit(commit, "Append Table row".into(), cx);
                true
            }
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
                false
            }
        })
    }

    pub(crate) fn add_table_row(&mut self, id: TableId, cx: &mut Context<Self>) {
        if self.mode.is_editing() || self.table_growth_blocked(cx) {
            return;
        }
        self.append_table_row_and_select(id, Vec::new(), cx);
    }

    fn append_table_row_and_select(
        &mut self,
        id: TableId,
        writes: Vec<(usize, usize, String)>,
        cx: &mut Context<Self>,
    ) {
        let result = self.workbook.update(cx, |wb, _| {
            if let [(row, col, value)] = writes.as_slice() {
                wb.append_table_rows_with_edit(id, 1, *row, *col, value)
            } else {
                wb.append_table_rows(id, 1, &writes)
            }
        });
        match result {
            Ok(commit) => {
                let range = commit.after_table().unwrap().range;
                let percent_cell = self
                    .mode
                    .is_editing()
                    .then_some(self.view_state.selected)
                    .filter(|_| self.edit_value.trim().ends_with('%'));
                self.cancel_edit(cx);
                self.record_table_commit(commit, "Append Table row".into(), cx);
                if let Some((row, col)) = percent_cell {
                    if matches!(
                        self.sheet(cx).get_format(row, col).number_format,
                        visigrid_engine::cell::NumberFormat::General
                    ) {
                        self.with_active_sheet_mut(cx, |sheet| {
                            sheet.set_number_format(
                                row,
                                col,
                                visigrid_engine::cell::NumberFormat::Percent { decimals: 0 },
                            )
                        });
                    }
                }
                self.clipboard_visual_range = None;
                self.maybe_show_cycle_banner(cx);
                self.surface_incremental_recalc_problems(cx);
                self.view_state.selected = (range.end_row, range.start_col);
                self.view_state.selection_end = None;
                self.view_state.additional_selections.clear();
                self.tab_chain_origin_col = Some(range.start_col);
                self.ensure_visible(cx);
            }
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
            }
        }
    }

    pub(crate) fn table_tab_append(&mut self, cx: &mut Context<Self>) -> bool {
        if self.is_multi_selection() || self.mode.is_overlay() {
            return false;
        }
        if self.mode.is_editing() {
            self.restore_formula_home_sheet(cx);
        }
        let Some(table) = self.table_under_cursor(cx) else {
            return false;
        };
        let (r, c) = self.view_state.selected;
        if table.range.data_rows() == 0
            || (self.row_view.view_to_data(r), c) != (table.range.end_row, table.range.end_col)
        {
            return false;
        }
        if self.table_growth_blocked(cx) {
            return true;
        }
        let writes = if self.mode.is_editing() {
            let mut value = self.edit_value.clone();
            if value.starts_with('+') {
                value = format!("={}", &value[1..]);
            }
            if value.starts_with('=') {
                let missing = value
                    .chars()
                    .filter(|c| *c == '(')
                    .count()
                    .saturating_sub(value.chars().filter(|c| *c == ')').count());
                value.extend(std::iter::repeat_n(')', missing));
            }
            vec![(r, c, value)]
        } else {
            Vec::new()
        };
        self.append_table_row_and_select(table.id, writes, cx);
        true
    }

    /// Growth consumes the paste in one guarded commit. Other pastes retain
    /// their existing clipboard behavior.
    pub(crate) fn paste_table_growth(
        &mut self,
        row: usize,
        col: usize,
        values: &[Vec<String>],
        incompatible_objects: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        let width = values.iter().map(Vec::len).max().unwrap_or(0);
        if values.is_empty() || width == 0 || self.sheet(cx).tables().is_empty() {
            return false;
        }
        let range = TableRange {
            start_row: row,
            start_col: col,
            end_row: row.saturating_add(values.len() - 1),
            end_col: col.saturating_add(width - 1),
        };
        let id = match self.wb(cx).table_append_target(self.sheet(cx).id, range) {
            Ok(Some(id)) => id,
            Ok(None) => return false,
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
                return true;
            }
        };
        if self.table_growth_blocked(cx) {
            return true;
        }
        if incompatible_objects {
            self.status_message = Some("This append includes clipboard merges or comments. Use Paste Values, or resize the Table before pasting.".into());
            cx.notify();
            return true;
        }
        let count = range.end_row - self.wb(cx).table(id).unwrap().1.range.end_row;
        let writes: Vec<_> = values
            .iter()
            .enumerate()
            .flat_map(|(r, cells)| {
                cells
                    .iter()
                    .enumerate()
                    .map(move |(c, value)| (row + r, col + c, value.clone()))
            })
            .collect();
        match self
            .workbook
            .update(cx, |wb, _| wb.append_table_rows(id, count, &writes))
        {
            Ok(commit) => {
                self.record_table_commit(
                    commit,
                    format!("Paste and append {count} Table row(s)"),
                    cx,
                );
                self.clipboard_visual_range = None;
                let failures = self.wb(cx).validate_range(
                    self.sheet_index(cx),
                    row,
                    col,
                    range.end_row,
                    range.end_col,
                );
                if failures.count > 0 {
                    self.store_validation_failures(&failures);
                    self.status_message = Some(format!(
                        "Pasted and appended {count} Table row(s); {} validation failure(s).",
                        failures.count
                    ));
                }
            }
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
            }
        }
        true
    }

    pub(crate) fn show_table_controls(&self, cx: &App) -> bool {
        let (row, col) = self.view_state.selected;
        !self.zen_mode
            && !self.is_previewing()
            && self.review_mode.is_none()
            && self
                .sheet(cx)
                .table_at(self.row_view.view_to_data(row), col)
                .is_some()
    }

    pub(crate) fn table_under_cursor(&self, cx: &App) -> Option<DataTable> {
        let (r, c) = self.view_state.selected;
        self.sheet(cx)
            .table_at(self.row_view.view_to_data(r), c)
            .cloned()
    }

    pub(crate) fn create_table_dialog(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) || self.mode.is_editing() || self.mode.is_overlay() {
            return;
        }
        if self.row_view.is_sorted() || self.row_view.is_filtered() {
            self.status_message = Some("Clear sorting and filters before creating a Table.".into());
            cx.notify();
            return;
        }
        if self.all_selection_ranges().len() > 1 {
            self.status_message = Some("Select one rectangle to create a Table.".into());
            cx.notify();
            return;
        }
        if self.table_under_cursor(cx).is_some() {
            self.open_table_dialog(
                TableDialogKind::Resize(self.table_under_cursor(cx).unwrap().id),
                cx,
            );
            return;
        }
        let ((r0, c0), (r1, c1)) = self.selection_range();
        let (r0, c0, r1, c1) = if r0 == r1 && c0 == c1 {
            crate::ai::find_current_region(self.sheet(cx), r0, c0)
        } else {
            (r0, c0, r1, c1)
        };
        self.table_dialog = Some(TableDialog {
            kind: TableDialogKind::Create,
            sheet: self.sheet(cx).id,
            name: self.wb(cx).next_table_name(),
            range: range_label(TableRange {
                start_row: r0,
                start_col: c0,
                end_row: r1,
                end_col: c1,
            }),
            has_headers: true,
            field: 0,
            select_all: true,
            error: None,
        });
        self.pivot_panel = None;
        cx.notify();
    }

    pub(crate) fn open_table_dialog(&mut self, kind: TableDialogKind, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) || self.mode.is_editing() || self.mode.is_overlay() {
            return;
        }
        let id = match kind {
            TableDialogKind::Rename(id)
            | TableDialogKind::Resize(id)
            | TableDialogKind::Convert(id)
            | TableDialogKind::ColumnFormula(id, _, _) => id,
            _ => return,
        };
        let Some((sheet, table)) = self.wb(cx).table(id) else {
            return;
        };
        self.table_dialog = Some(TableDialog {
            kind,
            sheet,
            name: if let TableDialogKind::ColumnFormula(_, col, _) = kind {
                let row = self.row_view.view_to_data(self.view_state.selected.0);
                if table.columns[col - table.range.start_col].formula.is_some() {
                    table.columns[col - table.range.start_col]
                        .formula
                        .clone()
                        .unwrap()
                } else {
                    self.sheet(cx).get_raw(row, col)
                }
            } else {
                table.name.clone()
            },
            range: if let TableDialogKind::ColumnFormula(_, col, _) = kind {
                let column = &table.columns[col - table.range.start_col];
                if column.formula.is_some() {
                    (table.range.start_row + column.formula_origin).to_string()
                } else {
                    self.row_view
                        .view_to_data(self.view_state.selected.0)
                        .to_string()
                }
            } else {
                range_label(table.range)
            },
            has_headers: true,
            field: usize::from(matches!(kind, TableDialogKind::Resize(_))),
            select_all: true,
            error: None,
        });
        self.pivot_panel = None;
        cx.notify();
    }

    pub(crate) fn submit_table_dialog(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) {
            return;
        }
        let Some(draft) = self.table_dialog.clone() else {
            return;
        };
        if draft.kind == TableDialogKind::Create && !draft.has_headers {
            let last = self.wb(cx).sheet_by_id(draft.sheet).map(|s| s.rows - 1).unwrap_or(crate::app::NUM_ROWS - 1);
            if self.row_heights.get(&draft.sheet).is_some_and(|h| h.contains_key(&last))
                || self.hidden_rows.get(&draft.sheet).is_some_and(|h| h.contains(&last)) {
                self.table_dialog.as_mut().unwrap().error = Some("Inserting a header would push row formatting off the sheet.".into());
                cx.notify(); return;
            }
        }
        let result = self.workbook.update(cx, |wb, _| match draft.kind {
            TableDialogKind::Create => parse_range(&draft.range)
                .and_then(|r| if draft.has_headers { wb.create_table(draft.sheet, r, draft.name.trim()) }
                    else { wb.create_table_without_headers(draft.sheet, r, draft.name.trim()) }),
            TableDialogKind::Rename(id) => wb.rename_table(id, draft.name.trim()),
            TableDialogKind::Resize(id) => {
                parse_range(&draft.range).and_then(|r| wb.resize_table(id, r))
            }
            TableDialogKind::Convert(id) => wb.remove_table(id),
            TableDialogKind::ColumnFormula(id, col, replace) => {
                let row = draft
                    .range
                    .parse::<usize>()
                    .map_err(|_| "Invalid formula origin.".to_string());
                row.and_then(|row| wb.set_calculated_column(id, col, row, &draft.name, replace))
            }
        });
        match result {
            Ok(commit) => {
                let verb = match draft.kind {
                    TableDialogKind::Create => "Create Table",
                    TableDialogKind::Rename(_) => "Rename Table",
                    TableDialogKind::Resize(_) => "Resize Table",
                    TableDialogKind::Convert(_) => "Convert Table to range",
                    TableDialogKind::ColumnFormula(_, _, _) => "Set column formula",
                };
                self.record_table_commit(commit, format!("{verb}: {}", draft.name.trim()), cx);
                self.table_dialog = None;
            }
            Err(error) => {
                if let Some(d) = self.table_dialog.as_mut() {
                    d.error = Some(error);
                }
            }
        }
        cx.notify();
    }

    pub(crate) fn toggle_table_banding(&mut self, id: TableId, cx: &mut Context<Self>) {
        if self.block_if_previewing(cx) || self.mode.is_editing() {
            return;
        }
        let Some((_, table)) = self.wb(cx).table(id) else {
            return;
        };
        let mut style = table.style.clone();
        style.banded_rows = !style.banded_rows;
        let result = self
            .workbook
            .update(cx, |wb, _| wb.set_table_style(id, style));
        match result {
            Ok(commit) => self.record_table_commit(commit, "Change Table banding".into(), cx),
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
            }
        }
    }

    pub(crate) fn record_table_commit(
        &mut self,
        commit: TableCommit,
        description: String,
        cx: &mut Context<Self>,
    ) {
        self.update_header_insertion_view(&commit, false, cx);
        self.history.record_action_with_provenance(
            UndoAction::TableCommit {
                sheet_index: self
                    .wb(cx)
                    .sheet_index_by_id(commit.sheet_id())
                    .unwrap_or(self.sheet_index(cx)),
                commit: Box::new(commit),
                description: description.clone(),
            },
            None,
        );
        self.bump_cells_rev();
        self.is_modified = true;
        self.status_message = Some(description);
        cx.notify();
    }

    pub(crate) fn replay_table_commit(
        &mut self,
        commit: &TableCommit,
        undo: bool,
        cx: &mut Context<Self>,
    ) -> bool {
        if crate::table_header_paste::is_header_rename(commit) {
            return self.replay_table_headers(commit, undo, cx);
        }
        match self
            .workbook
            .update(cx, |wb, _| wb.apply_table_commit(commit, undo))
        {
            Ok(()) => {
                self.update_header_insertion_view(commit, undo, cx);
                self.bump_cells_rev();
                self.is_modified = true;
                cx.notify();
                true
            }
            Err(e) => {
                self.status_message = Some(format!(
                    "Cannot {} Table change: {e}",
                    if undo { "undo" } else { "redo" }
                ));
                cx.notify();
                false
            }
        }
    }

    /// Mirror the composite header insertion in GUI-owned row presentation.
    fn update_header_insertion_view(&mut self, commit: &TableCommit, undo: bool, cx: &mut Context<Self>) {
        let Some(at) = commit.inserted_header_row() else { return; };
        let sheet = commit.sheet_id();
        if let Some(heights) = self.row_heights.get_mut(&sheet) {
            *heights = heights.drain().filter_map(|(row, height)| {
                if undo && row == at { None }
                else { Some((if row >= at { if undo { row - 1 } else { row + 1 } } else { row }, height)) }
            }).collect();
        }
        if let Some(hidden) = self.hidden_rows.get_mut(&sheet) {
            *hidden = hidden.iter().copied().filter_map(|row| {
                if undo && row == at { None }
                else { Some(if row >= at { if undo { row - 1 } else { row + 1 } } else { row }) }
            }).collect();
        }
        if self.sheet(cx).id == sheet {
            if undo { self.row_view.delete_row(at); } else { self.row_view.insert_row(at); }
            let range = commit.after_table().unwrap().range;
            self.view_state.selected = (range.start_row, range.start_col);
            self.view_state.selection_end = Some((range.end_row - usize::from(undo), range.end_col));
            self.view_state.additional_selections.clear();
        }
    }

    /// Header renames preserve identities and rewrite dependent formulas.
    /// Returns Some(success) for a header, None for an ordinary cell.
    pub(crate) fn commit_table_header(
        &mut self,
        view_row: usize,
        col: usize,
        value: &str,
        cx: &mut Context<Self>,
    ) -> Option<bool> {
        let row = self.row_view.view_to_data(view_row);
        let table = self.sheet(cx).table_header_at(row, col)?.clone();
        let mut names: Vec<_> = table.columns.iter().map(|c| c.name.clone()).collect();
        names[col - table.range.start_col] = value.to_string();
        let result = self
            .workbook
            .update(cx, |wb, _| wb.rename_table_columns(table.id, &names));
        Some(match result {
            Ok(commit) => {
                self.record_table_commit(commit, format!("Rename {} column", table.name), cx);
                true
            }
            Err(e) => {
                self.status_message = Some(e);
                cx.notify();
                false
            }
        })
    }

    pub(crate) fn block_if_table_header(
        &mut self,
        r0: usize,
        c0: usize,
        r1: usize,
        c1: usize,
        what: &str,
        cx: &mut Context<Self>,
    ) -> bool {
        let sheet = self.sheet(cx);
        let name = sheet
            .tables()
            .iter()
            .find(|t| header_in_view_rect(t, &self.row_view, r0, c0, r1, c1))
            .map(|t| t.name.clone());
        if let Some(name) = name {
            self.status_message = Some(format!(
                "Cannot {what} across {name}'s headers. Edit a header cell to rename its column."
            ));
            cx.notify();
            true
        } else {
            false
        }
    }

    pub(crate) fn block_selection_table_headers(
        &mut self,
        what: &str,
        cx: &mut Context<Self>,
    ) -> bool {
        for ((r0, c0), (r1, c1)) in self.all_selection_ranges() {
            if self.block_if_table_header(r0, c0, r1, c1, what, cx) {
                return true;
            }
        }
        false
    }

    pub(crate) fn block_table_paste(
        &mut self,
        row: usize,
        col: usize,
        rows: usize,
        cols: usize,
        cx: &mut Context<Self>,
    ) -> bool {
        if rows == 0 || cols == 0 || self.sheet(cx).tables().is_empty() {
            return false;
        }
        if rows > 1 && (self.row_view.is_sorted() || self.row_view.is_filtered()) {
            self.status_message =
                Some("Clear sorting and filters before pasting multiple Table rows.".into());
            cx.notify();
            return true;
        }
        let start = self.row_view.visible_rows().iter().position(|r| *r == row);
        for offset in 0..rows {
            let view_row = if self.row_view.is_filtered() {
                let Some(r) = start.and_then(|s| self.row_view.nth_visible(s + offset)) else {
                    continue;
                };
                r
            } else {
                row + offset
            };
            if self.block_if_table_header(view_row, col, view_row, col + cols - 1, "paste", cx) {
                return true;
            }
        }
        false
    }

    /// A modal interceptor runs before grid bindings so typing, Enter, Delete,
    /// paste and Ctrl+T cannot mutate cells behind the dialog.
    pub(crate) fn table_dialog_key(&mut self, key: &Keystroke, cx: &mut Context<Self>) -> bool {
        use crate::ui::text_input::{handle_input_key, handle_input_paste, InputAction};
        let Some(d) = self.table_dialog.as_mut() else {
            return false;
        };
        if key.key == "escape" {
            self.table_dialog = None;
            cx.notify();
            return true;
        }
        if key.key == "enter" {
            self.submit_table_dialog(cx);
            return true;
        }
        if matches!(d.kind, TableDialogKind::Convert(_)) {
            return true;
        }
        if key.key == "tab" && d.kind == TableDialogKind::Create {
            d.field = (d.field + if key.modifiers.shift { 2 } else { 1 }) % 3;
            d.select_all = true;
            cx.notify();
            return true;
        }
        if d.field == 2 {
            if key.key == "space" || key.key_char.as_deref() == Some(" ") {
                d.has_headers = !d.has_headers;
                d.error = None;
                cx.notify();
            }
            return true;
        }
        let buffer = if d.field == 0 {
            &mut d.name
        } else {
            &mut d.range
        };
        if key.modifiers.control || key.modifiers.platform {
            match key.key.as_str() {
                "a" => d.select_all = true,
                "c" => {
                    if d.select_all {
                        cx.write_to_clipboard(ClipboardItem::new_string(buffer.clone()));
                    }
                }
                "x" => {
                    if d.select_all {
                        cx.write_to_clipboard(ClipboardItem::new_string(buffer.clone()));
                        buffer.clear();
                        d.select_all = false;
                    }
                }
                "v" => {
                    if let Some(text) = cx.read_from_clipboard().and_then(|c| c.text()) {
                        handle_input_paste(buffer, &mut d.select_all, &text);
                    }
                }
                _ => {}
            }
        } else if handle_input_key(
            buffer,
            &mut d.select_all,
            &key.key,
            key.key_char.as_deref(),
            key.modifiers.alt,
        ) == InputAction::Changed
        {
            d.error = None;
        }
        cx.notify();
        true
    }
}

#[cfg(test)]
mod tests {
    use super::{header_in_view_rect, parse_range, range_label, TableRange};
    #[test]
    fn table_dialog_ranges_are_finite_local_and_ordered() {
        assert_eq!(range_label(parse_range("$B$2:D20").unwrap()), "B2:D20");
        for s in [
            "D20:B2",
            "A:A",
            "1:4",
            "Other!A1:B3",
            "Sales",
            "A0",
            "XFE1",
            "A1048577",
        ] {
            assert!(parse_range(s).is_err(), "{s}");
        }
    }
    #[test]
    fn header_preflight_uses_destinations_and_canonical_rows() {
        let mut wb = visigrid_engine::workbook::Workbook::new();
        let sheet = wb.active_sheet_id();
        let commit = wb
            .create_table(
                sheet,
                TableRange {
                    start_row: 2,
                    start_col: 2,
                    end_row: 5,
                    end_col: 3,
                },
                "Sales",
            )
            .unwrap();
        let table = wb.table(commit.table_id()).unwrap().1;
        let mut rows = visigrid_engine::filter::RowView::new(10);
        assert!(header_in_view_rect(table, &rows, 0, 0, 9, 4));
        assert!(!header_in_view_rect(table, &rows, 3, 2, 5, 3));
        assert!(!header_in_view_rect(table, &rows, 0, 4, 9, 5));
        rows.apply_sort(vec![2, 1, 0, 3, 4, 5, 6, 7, 8, 9]);
        assert!(header_in_view_rect(table, &rows, 0, 2, 0, 3));
        assert!(!header_in_view_rect(table, &rows, 2, 2, 2, 3));
    }
}
