//! Pivot tables in the desktop app: commands, the field-list drawer state,
//! refresh, refusals for edits that touch pivot output, and undo/redo glue.
//!
//! The engine owns the rules (crates/engine/src/pivot.rs, workbook_pivot.rs):
//! ownership is enforced there, and every pivot action is a scoped
//! `PivotCommit`. This module decides when to run them and what to tell the
//! user.

use gpui::*;
use visigrid_engine::filter::RowView;
use visigrid_engine::sheet::Sheet;
use visigrid_engine::workbook::PivotCommit;

use crate::app::{Spreadsheet, NUM_ROWS};

/// Preflight and replay the complete pivot action, including sheet creation
/// or removal. The candidate is published only after every saved view rebuilds.
fn replay_pivot_candidate(
    wb: &visigrid_engine::workbook::Workbook,
    commit: &PivotCommit,
    created_sheet: &Option<(usize, Box<Sheet>)>,
    undo: bool,
) -> Result<visigrid_engine::workbook::Workbook, String> {
    let mut candidate = wb.clone();
    if !undo {
        if let Some((index, sheet)) = created_sheet {
            if candidate.sheet_index_by_id(sheet.id).is_none() && !candidate.restore_sheet(*index, (**sheet).clone()) {
                return Err("The pivot's sheet could not be restored.".into());
            }
        }
    }
    candidate.apply_pivot_state(if undo { &commit.before } else { &commit.after }).map_err(|e| e.to_string())?;
    if undo {
        if let Some((_, sheet)) = created_sheet {
            if let Some(index) = candidate.sheet_index_by_id(sheet.id) {
                candidate.take_sheet(index).ok_or("The pivot's sheet could not be removed.")?;
                let report = candidate.recompute_full_ordered();
                if let Some(error) = report.errors.first() {
                    return Err(format!("The pivot could not be recalculated: {error:?}"));
                }
            }
        }
    }
    for sheet in candidate.sheets() { sheet.build_saved_table_view(NUM_ROWS.min(sheet.rows))?; }
    if let Some(error) = candidate.take_incremental_errors().first() {
        return Err(format!("The pivot could not be recalculated: {error:?}"));
    }
    Ok(candidate)
}

/// Only the exact pivot-creation group (pivot + its new sheet's widths) may
/// bypass the general unrelated-history gate while Table criteria are active.
pub(crate) fn is_pivot_history(action: &crate::history::UndoAction) -> bool {
    use crate::history::UndoAction;
    match action {
        UndoAction::PivotCommit { .. } => true,
        UndoAction::Group { actions, .. } => {
            let Some(UndoAction::PivotCommit { created_sheet: Some((_, sheet)), .. }) = actions.first() else { return false };
            actions.iter().skip(1).all(|a| matches!(a, UndoAction::ColumnWidthSet { sheet_id, .. } if *sheet_id == sheet.id))
        }
        _ => false,
    }
}

impl Spreadsheet {
    /// Undo a pivot action: restore the "before" state, and remove the sheet
    /// the action created, if any.
    pub(crate) fn pivot_undo(
        &mut self,
        commit: &PivotCommit,
        created_sheet: &Option<(usize, Box<Sheet>)>,
        cx: &mut Context<Self>,
    ) {
        let active_before = self.wb(cx).active_sheet_index();
        let candidate = match replay_pivot_candidate(self.wb(cx), commit, created_sheet, true) {
            Ok(candidate) => candidate,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        self.workbook.update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.finish_pivot_restore(active_before, cx);
    }

    /// Redo a pivot action: restore the sheet it created, if any, then apply
    /// the "after" state.
    pub(crate) fn pivot_redo(
        &mut self,
        commit: &PivotCommit,
        created_sheet: &Option<(usize, Box<Sheet>)>,
        cx: &mut Context<Self>,
    ) {
        let active_before = self.wb(cx).active_sheet_index();
        let candidate = match replay_pivot_candidate(self.wb(cx), commit, created_sheet, false) {
            Ok(candidate) => candidate,
            Err(error) => { self.status_message = Some(error); cx.notify(); return; }
        };
        self.workbook.update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
        self.finish_pivot_restore(active_before, cx);
    }

    pub(crate) fn preflight_pivot_history(&self, action: &crate::history::UndoAction, undo: bool, cx: &App) -> Result<(), String> {
        use crate::history::UndoAction;
        let pivot = match action {
            UndoAction::Group { actions, .. } => actions.first(),
            action => Some(action),
        };
        if let Some(UndoAction::PivotCommit { commit, created_sheet, .. }) = pivot {
            replay_pivot_candidate(self.wb(cx), commit, created_sheet, undo)?;
        }
        Ok(())
    }

    fn finish_pivot_restore(&mut self, active_before: usize, cx: &mut Context<Self>) {
        let active_after = self.wb(cx).active_sheet_index();
        let row_view = if active_after == active_before {
            self.row_view.clone()
        } else {
            RowView::new(NUM_ROWS)
        };
        self.finish_workbook_snapshot_restore(row_view, cx);
        self.sync_table_view(cx);
        self.is_modified = true;
        cx.notify();
    }
}

// ---------------------------------------------------------------------------
// Refusals: an edit that touches pivot output is refused as a whole.
// ---------------------------------------------------------------------------

impl Spreadsheet {
    /// The pivot whose owned output intersects the rectangle given in view
    /// coordinates (rows as displayed, so sort/filter are mapped) on the active
    /// sheet. Returns its name.
    pub(crate) fn pivot_in_view_rect(
        &self,
        view_r0: usize,
        c0: usize,
        view_r1: usize,
        c1: usize,
        cx: &App,
    ) -> Option<String> {
        let sheet = self.sheet(cx);
        if sheet.pivots.is_empty() {
            return None;
        }
        let (r0, r1) = (view_r0.min(view_r1), view_r0.max(view_r1));
        let (c0, c1) = (c0.min(c1), c0.max(c1));
        if !self.row_view.is_sorted() && !self.row_view.is_filtered() {
            return sheet.pivot_in_rect(r0, c0, r1, c1).map(|p| p.name.clone());
        }
        // Sorted or filtered: map each displayed row to its data row.
        sheet.pivots.iter().find_map(|p| {
            let (pr0, pc0, pr1, pc1) = p.region()?;
            if pc0 > c1 || c0 > pc1 {
                return None;
            }
            (r0..=r1)
                .map(|vr| self.row_view.view_to_data(vr))
                .any(|dr| dr >= pr0 && dr <= pr1)
                .then(|| p.name.clone())
        })
    }

    /// Refuse an edit whose target touches pivot output: set a status message
    /// and return true. `what` names the operation ("edit", "paste", …).
    pub(crate) fn block_if_pivot(
        &mut self,
        view_r0: usize,
        c0: usize,
        view_r1: usize,
        c1: usize,
        what: &str,
        cx: &mut Context<Self>,
    ) -> bool {
        let Some(name) = self.pivot_in_view_rect(view_r0, c0, view_r1, c1, cx) else {
            return false;
        };
        self.status_message = Some(format!(
            "Can't {what} here: this is {name}'s output. Change it from the field list, or copy the values to another sheet."
        ));
        cx.notify();
        true
    }

    /// Refuse if any of the current selection ranges touches pivot output.
    pub(crate) fn block_if_selection_in_pivot(&mut self, what: &str, cx: &mut Context<Self>) -> bool {
        let ranges = self.all_selection_ranges();
        for ((r0, c0), (r1, c1)) in ranges {
            if self.block_if_pivot(r0, c0, r1, c1, what, cx) {
                return true;
            }
        }
        false
    }
}

// ---------------------------------------------------------------------------
// Field-list drawer state, and the create / apply / refresh / delete flow.
// ---------------------------------------------------------------------------

use visigrid_engine::cell::NumberFormat;
use visigrid_engine::pivot::{
    self, Aggregation, PivotDefinition, PivotError, PivotField, PivotOutput, PivotSnapshot, PivotSource,
    PivotTable, PivotValueField,
};

/// Sources with more data rows than this aggregate on a background thread.
const BACKGROUND_ROWS: u32 = 100_000;

use visigrid_engine::pivot::{format_new_pivot_values, style_new_pivot};

/// What the drawer is building.
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum PivotPanelMode {
    /// A new pivot, not yet applied. No output sheet exists until Apply.
    New,
    /// Editing an existing pivot.
    Edit { pivot_id: u64 },
}

/// One line of the drawer's keyboard list.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum PivotPanelItem {
    Field(usize),
    Row(usize),
    Column,
    Value(usize),
}

/// The field-list drawer. Edits a draft definition; nothing changes in the
/// workbook until Apply.
#[derive(Debug, Clone)]
pub(crate) struct PivotPanel {
    pub mode: PivotPanelMode,
    pub source: PivotSource,
    pub headers: Vec<String>,
    pub column_ids: Vec<Option<visigrid_engine::table::TableColumnId>>,
    pub table_name: Option<String>,
    pub source_menu: bool,
    pub source_cursor: usize,
    /// Per source column: first non-empty data cell's number format, and
    /// whether it holds a number (drives the default aggregation).
    pub column_formats: Vec<NumberFormat>,
    pub column_numeric: Vec<bool>,
    pub draft: PivotDefinition,
    pub cursor: usize,
    pub message: Option<String>,
    /// Last row of data appended below the source, if detected.
    pub growth: Option<u32>,
    pub busy: bool,
}

impl PivotPanel {
    /// The keyboard list: every source field, then the Rows, Column and
    /// Values wells.
    pub fn items(&self) -> Vec<PivotPanelItem> {
        let mut v: Vec<PivotPanelItem> = (0..self.headers.len()).map(PivotPanelItem::Field).collect();
        v.extend((0..self.draft.rows.len()).map(PivotPanelItem::Row));
        if self.draft.column.is_some() {
            v.push(PivotPanelItem::Column);
        }
        v.extend((0..self.draft.values.len()).map(PivotPanelItem::Value));
        v
    }

    pub fn current(&self) -> Option<PivotPanelItem> {
        self.items().get(self.cursor).copied()
    }

    fn field(&self, offset: usize) -> PivotField {
        PivotField { column_id: self.column_ids.get(offset).copied().flatten(), offset: offset as u32, header: self.headers[offset].clone() }
    }

    pub(crate) fn move_cursor_to(&mut self, item: PivotPanelItem) {
        if let Some(i) = self.items().iter().position(|x| *x == item) {
            self.cursor = i;
        }
    }

    pub fn source_label(&self) -> String {
        format!(
            "{}{}:{}{}",
            Spreadsheet::col_letter(self.source.start_col as usize),
            self.source.start_row + 1,
            Spreadsheet::col_letter(self.source.end_col as usize),
            self.source.end_row + 1
        )
    }

    /// Assign the field under the cursor. Returns a message if nothing happened.
    pub(crate) fn assign(&mut self, target: char) -> Option<String> {
        let Some(PivotPanelItem::Field(i)) = self.current() else {
            return Some("Move to a field in the list first.".into());
        };
        let f = self.field(i);
        match target {
            'r' => {
                if self.draft.rows.iter().any(|x| x.offset == f.offset)
                    || self.draft.column.as_ref().is_some_and(|c| c.offset == f.offset)
                {
                    return Some(format!("{} is already a row or column field.", f.header));
                }
                self.draft.rows.push(f);
            }
            'c' => {
                if self.draft.rows.iter().any(|x| x.offset == f.offset) {
                    return Some(format!("{} is already a row field.", f.header));
                }
                self.draft.column = Some(f);
            }
            'v' => {
                let aggregation = if self.column_numeric.get(i).copied().unwrap_or(false) {
                    Aggregation::Sum
                } else {
                    Aggregation::Count
                };
                let number_format =
                    pivot::default_number_format(aggregation, self.column_formats.get(i).unwrap_or(&NumberFormat::General));
                self.draft.values.push(PivotValueField { field: f, aggregation, number_format });
            }
            _ => {}
        }
        None
    }

    pub(crate) fn remove_current(&mut self) {
        match self.current() {
            Some(PivotPanelItem::Row(k)) => {
                self.draft.rows.remove(k);
            }
            Some(PivotPanelItem::Column) => self.draft.column = None,
            Some(PivotPanelItem::Value(k)) => {
                self.draft.values.remove(k);
            }
            _ => return,
        }
        self.cursor = self.cursor.min(self.items().len().saturating_sub(1));
    }

    pub(crate) fn reorder(&mut self, up: bool) {
        match self.current() {
            Some(PivotPanelItem::Row(k)) => {
                let j = if up { k.checked_sub(1) } else { (k + 1 < self.draft.rows.len()).then_some(k + 1) };
                if let Some(j) = j {
                    self.draft.rows.swap(k, j);
                    self.move_cursor_to(PivotPanelItem::Row(j));
                }
            }
            Some(PivotPanelItem::Value(k)) => {
                let j = if up { k.checked_sub(1) } else { (k + 1 < self.draft.values.len()).then_some(k + 1) };
                if let Some(j) = j {
                    self.draft.values.swap(k, j);
                    self.move_cursor_to(PivotPanelItem::Value(j));
                }
            }
            _ => {}
        }
    }

    /// Cycle the aggregation of the value field under the cursor. The number
    /// format follows: counts become whole numbers, other aggregations take
    /// the source column's format again.
    fn cycle_aggregation(&mut self, forward: bool) {
        let Some(PivotPanelItem::Value(k)) = self.current() else { return };
        let all = Aggregation::ALL;
        let cur = all.iter().position(|a| *a == self.draft.values[k].aggregation).unwrap_or(0);
        let next = if forward { (cur + 1) % all.len() } else { (cur + all.len() - 1) % all.len() };
        self.set_aggregation(k, all[next]);
    }

    pub(crate) fn set_aggregation(&mut self, k: usize, agg: Aggregation) {
        let offset = self.draft.values[k].field.offset as usize;
        let src = self.column_formats.get(offset).cloned().unwrap_or_default();
        let v = &mut self.draft.values[k];
        let was_count = v.aggregation.is_count();
        v.aggregation = agg;
        // Re-derive the format only when it was a default (count ↔ value).
        if was_count != agg.is_count() {
            v.number_format = pivot::default_number_format(agg, &src);
        }
    }

    /// Apply the aggregation under the cursor to every value field.
    pub(crate) fn apply_aggregation_to_all(&mut self) -> Option<String> {
        let Some(PivotPanelItem::Value(k)) = self.current() else {
            return Some("Move to a value field first.".into());
        };
        let agg = self.draft.values[k].aggregation;
        for i in 0..self.draft.values.len() {
            self.set_aggregation(i, agg);
        }
        Some(format!("All value fields now use {}.", agg.label()))
    }
}

/// A computation in flight: what to place once the output is ready.
struct PivotJob {
    mode: PivotPanelMode,
    /// Show the new sheet when a create finishes. False for session clients:
    /// an agent's pivot must not move the user's view or drop their typing.
    activate: bool,
    table: PivotTable,
    source_generation: u64,
    description: String,
}

impl Spreadsheet {
    // ---- opening the drawer -------------------------------------------------

    /// Insert → PivotTable: the selection (if more than one cell) or the
    /// current region around the cursor becomes the source.
    pub(crate) fn insert_pivot_table(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) {
            return;
        }
        let (vr, col) = self.view_state.selected;
        let row = self.row_view.view_to_data(vr);
        if let Some(table) = self.sheet(cx).table_at(row, col) {
            let source = self.wb(cx).table_pivot_source(table.id).unwrap();
            self.open_pivot_source(PivotPanelMode::New, source, cx);
            return;
        }
        if self.row_view.is_sorted() || self.row_view.is_filtered() {
            self.status_message = Some("Clear sorting and filters before creating a pivot table.".into());
            cx.notify();
            return;
        }
        let ((r0, c0), (r1, c1)) = self.selection_range();
        let (r0, c0, r1, c1) = if r0 == r1 && c0 == c1 {
            crate::ai::find_current_region(self.sheet(cx), r0, c0)
        } else {
            (r0, c0, r1, c1)
        };
        if r1 <= r0 {
            let source = self.wb(cx).tables().next().and_then(|(_, t)| self.wb(cx).table_pivot_source(t.id).ok());
            if let Some(source) = source {
                self.open_pivot_source(PivotPanelMode::New, source, cx);
                return;
            }
            self.status_message = Some("A pivot needs a header row and at least one data row. Select the table first.".into());
            cx.notify();
            return;
        }
        if let Some(name) = self.pivot_in_view_rect(r0, c0, r1, c1, cx) {
            self.status_message = Some(format!("The selection includes {name}'s output. Select the source data instead."));
            cx.notify();
            return;
        }
        let source = PivotSource {
            table_id: None, sheet_id: self.sheet(cx).id,
            start_row: r0 as u32,
            start_col: c0 as u32,
            end_row: r1 as u32,
            end_col: c1 as u32,
        };
        match self.build_pivot_panel(PivotPanelMode::New, source, PivotDefinition::default(), cx) {
            Ok(panel) => {
                self.pivot_panel = Some(panel);
                self.cf_panel_visible = false;
            }
            Err(msg) => self.status_message = Some(msg),
        }
        cx.notify();
    }

    fn open_pivot_source(&mut self, mode: PivotPanelMode, source: PivotSource, cx: &mut Context<Self>) {
        match self.build_pivot_panel(mode, source, PivotDefinition::default(), cx) {
            Ok(panel) => { self.pivot_panel = Some(panel); self.cf_panel_visible = false; }
            Err(error) => self.status_message = Some(error),
        }
        cx.notify();
    }

    pub(crate) fn pivot_choose_table(&mut self, id: visigrid_engine::table::TableId, cx: &mut Context<Self>) {
        let Some(panel) = self.pivot_panel.as_ref().filter(|p| !p.busy) else { return };
        if panel.source.table_id == Some(id) {
            self.pivot_panel.as_mut().unwrap().source_menu = false;
            cx.notify();
            return;
        }
        let mode = panel.mode.clone();
        match self.wb(cx).table_pivot_source(id) {
            Ok(source) => self.open_pivot_source(mode, source, cx),
            Err(error) => { self.set_pivot_panel_message(error.to_string()); cx.notify(); }
        }
    }

    /// Open the field list for the pivot under the cursor.
    pub(crate) fn edit_pivot_fields(&mut self, cx: &mut Context<Self>) {
        let Some(table) = self.pivot_under_cursor(cx) else {
            self.status_message = Some("Put the cursor in a pivot table to edit its fields.".into());
            cx.notify();
            return;
        };
        match self.build_pivot_panel(PivotPanelMode::Edit { pivot_id: table.id }, table.source, table.definition.clone(), cx) {
            Ok(mut panel) => {
                panel.growth = self.wb(cx).pivot_source_growth(&table);
                self.pivot_panel = Some(panel);
                self.cf_panel_visible = false;
            }
            Err(msg) => self.status_message = Some(msg),
        }
        cx.notify();
    }

    fn build_pivot_panel(
        &self,
        mode: PivotPanelMode,
        mut source: PivotSource,
        mut draft: PivotDefinition,
        cx: &App,
    ) -> Result<PivotPanel, String> {
        let wb = self.wb(cx);
        let mut message = None;
        let (table_name, column_ids) = if let Some(id) = source.table_id {
            source = match wb.table_pivot_source(id) {
                Ok(source) => source,
                Err(error) => return Ok(PivotPanel { mode, source, draft, headers: Vec::new(),
                    column_ids: Vec::new(), column_formats: Vec::new(), column_numeric: Vec::new(),
                    table_name: Some("Missing Table".into()), source_menu: true, source_cursor: 0,
                    cursor: 0, message: Some(error.to_string()), growth: None, busy: false }),
            };
            let (_, table) = wb.table(id).unwrap();
            // Keep missing fields visible and removable in the drawer. Apply
            // refuses them until the user repairs the definition.
            for field in draft.rows.iter_mut().chain(draft.column.iter_mut()).chain(draft.values.iter_mut().map(|v| &mut v.field)) {
                if let Some((offset, column)) = table.columns.iter().enumerate().find(|(_, c)| Some(c.id) == field.column_id) {
                    field.offset = offset as u32;
                    field.header = column.name.clone();
                } else {
                    field.offset = u32::MAX;
                    message = Some(format!("Source field '{}' is missing. Remove it from the layout and choose a replacement.", field.header));
                }
            }
            (Some(table.name.clone()), table.columns.iter().map(|c| Some(c.id)).collect())
        } else { (None, Vec::new()) };
        let sheet = wb.sheet_by_id(source.sheet_id).ok_or("The pivot's source sheet no longer exists.")?;
        let hr = source.start_row as usize;
        let headers: Vec<String> = (source.start_col..=source.end_col)
            .map(|c| sheet.get_display(hr, c as usize).trim().to_string())
            .collect();
        pivot::validate_headers(&headers).map_err(|e| e.to_string())?;
        // Shared with headless and agent-created pivots, so their defaults match.
        let (column_formats, column_numeric) =
            pivot::column_profile(sheet, &source).into_iter().map(|p| (p.format, p.numeric)).unzip();
        Ok(PivotPanel {
            mode,
            source,
            headers,
            column_ids,
            table_name,
            source_menu: false, source_cursor: 0,
            column_formats,
            column_numeric,
            draft,
            cursor: 0,
            message,
            growth: None,
            busy: false,
        })
    }

    /// The pivot whose output contains the cursor, on the active sheet.
    pub(crate) fn pivot_under_cursor(&self, cx: &App) -> Option<PivotTable> {
        let (vr, c) = self.view_state.selected;
        let r = self.row_view.view_to_data(vr);
        self.sheet(cx).pivot_at(r, c).cloned()
    }

    pub(crate) fn close_pivot_panel(&mut self, cx: &mut Context<Self>) {
        self.pivot_panel = None;
        cx.notify();
    }

    // ---- keyboard ---------------------------------------------------------------

    /// Keys for the open drawer. Returns true if the key was handled.
    /// Give the open field list first claim on keystrokes.
    ///
    /// GPUI dispatches key bindings before key-down listeners, and the grid
    /// binds the keys the list needs (arrows to MoveUp/Down/Left/Right, Enter
    /// to ConfirmEdit, Escape, Home/End, Delete to DeleteCell). A listener
    /// never saw them, so arrows moved the grid cursor and Delete cleared a
    /// cell. Interceptors run before bindings. This one acts only for this
    /// window, only while the grid itself has focus (not a dialog, palette,
    /// terminal or text field), with no menu open and no cell being edited.
    pub(crate) fn intercept_pivot_keys(window: &mut Window, cx: &mut Context<Self>) -> gpui::Subscription {
        let this = cx.entity().downgrade();
        let window_handle = window.window_handle();
        cx.intercept_keystrokes(move |event, window, cx| {
            if window.window_handle() != window_handle {
                return;
            }
            let Some(this) = this.upgrade() else { return };
            let handled = this.update(cx, |this, cx| {
                if this.table_filter_key(&event.keystroke, cx) { return true; }
                if this.table_dialog_key(&event.keystroke, cx) { return true; }
                if this.pivot_panel.is_none()
                    || this.open_menu.is_some()
                    || this.mode.is_editing()
                    || this.mode.is_overlay()
                    || !this.focus_handle.is_focused(window)
                {
                    return false;
                }
                this.pivot_panel_handle_key(&event.keystroke, cx)
            });
            if handled {
                cx.stop_propagation();
            }
        })
    }

    pub(crate) fn pivot_panel_handle_key(&mut self, keystroke: &gpui::Keystroke, cx: &mut Context<Self>) -> bool {
        let table_ids: Vec<_> = self.wb(cx).tables().map(|(_, t)| t.id).collect();
        let Some(panel) = self.pivot_panel.as_mut() else { return false };
        let key = keystroke.key.as_str();
        let m = &keystroke.modifiers;
        if panel.busy && key != "escape" {
            return true;
        }
        if panel.source_menu {
            match key {
                "escape" | "s" => panel.source_menu = false,
                "up" => panel.source_cursor = panel.source_cursor.saturating_sub(1),
                "down" => panel.source_cursor = (panel.source_cursor + 1).min(table_ids.len().saturating_sub(1)),
                "enter" => { if let Some(id) = table_ids.get(panel.source_cursor).copied() { self.pivot_choose_table(id, cx); } }
                _ => {},
            }
            cx.notify();
            return true;
        }
        let n = panel.items().len();
        let mut apply = false;
        match key {
            "escape" => {
                self.pivot_panel = None;
                cx.notify();
                return true;
            }
            "up" if m.alt => panel.reorder(true),
            "down" if m.alt => panel.reorder(false),
            "up" => panel.cursor = panel.cursor.saturating_sub(1),
            "down" => panel.cursor = (panel.cursor + 1).min(n.saturating_sub(1)),
            "home" => panel.cursor = 0,
            "end" => panel.cursor = n.saturating_sub(1),
            "left" => panel.cycle_aggregation(false),
            "right" => panel.cycle_aggregation(true),
            "delete" | "backspace" => panel.remove_current(),
            "enter" => apply = true,
            "f5" if m.alt => apply = true,
            _ if !m.control && !m.alt && !m.platform => match key {
                "s" => { panel.source_menu = true; panel.source_cursor = 0; }
                "r" | "c" | "v" => panel.message = panel.assign(key.chars().next().unwrap()),
                "=" => panel.message = panel.apply_aggregation_to_all(),
                "x" => {
                    if let Some(last) = panel.growth.take() {
                        let added = last - panel.source.end_row;
                        panel.source.end_row = last;
                        panel.message = Some(format!(
                            "Source extended by {added} row{} to {}. Apply to use it.",
                            if added == 1 { "" } else { "s" },
                            panel.source_label()
                        ));
                    } else {
                        panel.message = Some("No new rows were found below the source.".into());
                    }
                }
                _ => return false,
            },
            _ => return false,
        }
        if apply {
            self.pivot_apply(cx);
        }
        cx.notify();
        true
    }

    // ---- apply / refresh --------------------------------------------------------

    /// Apply the drawer's draft: create the pivot on a new sheet, or update
    /// the existing one.
    pub(crate) fn pivot_apply(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        let Some(panel) = self.pivot_panel.clone() else { return };
        let table = match &panel.mode {
            PivotPanelMode::New => PivotTable {
                id: self.wb(cx).next_pivot_id(),
                name: self.wb(cx).next_pivot_name(),
                source: panel.source,
                definition: panel.draft.clone(),
                anchor_row: 0,
                anchor_col: 0,
                extent: None,
                last_refresh: None,
                stale: false,
                source_generation: None,
            },
            PivotPanelMode::Edit { pivot_id } => match self.wb(cx).find_pivot(*pivot_id) {
                Some((_, t)) => PivotTable { source: panel.source, definition: panel.draft.clone(), ..t.clone() },
                None => {
                    self.set_pivot_panel_message("This pivot table no longer exists.".into());
                    return;
                }
            },
        };
        let description = match panel.mode {
            PivotPanelMode::New => format!("Create {}", table.name),
            PivotPanelMode::Edit { .. } => format!("Update {}", table.name),
        };
        self.run_pivot_job(PivotJob { mode: panel.mode.clone(), activate: true, table, source_generation: 0, description }, cx);
    }

    /// Refresh the pivot under the cursor (or the one in the open drawer).
    pub(crate) fn refresh_pivot(&mut self, cx: &mut Context<Self>) {
        let table = match self.pivot_panel.as_ref().map(|p| p.mode.clone()) {
            Some(PivotPanelMode::Edit { pivot_id }) => self.wb(cx).find_pivot(pivot_id).map(|(_, t)| t.clone()),
            _ => self.pivot_under_cursor(cx),
        };
        let Some(table) = table else {
            self.status_message = Some("Put the cursor in a pivot table to refresh it (Ctrl+Alt+F5 refreshes all).".into());
            cx.notify();
            return;
        };
        let description = format!("Refresh {}", table.name);
        let mode = PivotPanelMode::Edit { pivot_id: table.id };
        self.run_pivot_job(PivotJob { mode, activate: true, table, source_generation: 0, description }, cx);
    }

    /// Refresh every pivot in the workbook, one after another.
    pub(crate) fn refresh_all_pivots(&mut self, cx: &mut Context<Self>) {
        let tables: Vec<PivotTable> = self.wb(cx).pivots().into_iter().map(|(_, t)| t.clone()).collect();
        if tables.is_empty() {
            self.status_message = Some("This workbook has no pivot tables.".into());
            cx.notify();
            return;
        }
        let count = tables.len();
        let mut failed = 0;
        for table in tables {
            let description = format!("Refresh {}", table.name);
            let mode = PivotPanelMode::Edit { pivot_id: table.id };
            if !self.run_pivot_job_sync(PivotJob { mode, activate: true, table, source_generation: 0, description }, cx) {
                failed += 1;
            }
        }
        if failed == 0 {
            self.status_message = Some(format!("Refreshed {count} pivot table{}.", if count == 1 { "" } else { "s" }));
        } else {
            self.status_message = Some(format!("Refreshed {} of {count} pivot tables; see Problems for the rest.", count - failed));
        }
        cx.notify();
    }

    /// Delete the pivot under the cursor (or in the open drawer) and clear its
    /// output. One undo step.
    pub(crate) fn delete_pivot(&mut self, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        let id = match self.pivot_panel.as_ref().map(|p| p.mode.clone()) {
            Some(PivotPanelMode::Edit { pivot_id }) => Some(pivot_id),
            _ => self.pivot_under_cursor(cx).map(|t| t.id),
        };
        let Some(id) = id else {
            self.status_message = Some("Put the cursor in a pivot table to delete it.".into());
            cx.notify();
            return;
        };
        let commit = match self.wb(cx).prepare_pivot_delete(id) {
            Ok(c) => c,
            Err(e) => {
                self.status_message = Some(e.to_string());
                cx.notify();
                return;
            }
        };
        let name = commit.before.table.as_ref().map(|t| t.name.clone()).unwrap_or_default();
        if let Err(error) = self.workbook.update(cx, |wb, _| wb.apply_pivot_state(&commit.after)) {
            self.status_message = Some(error.to_string()); cx.notify(); return;
        }
        self.sync_table_view(cx);
        self.record_pivot_commit(commit, None, format!("Delete {name}"));
        self.pivot_panel = None;
        self.pivot_errors.remove(&id);
        self.status_message = Some(format!("Deleted {name}."));
        self.bump_cells_rev();
        self.is_modified = true;
        cx.notify();
    }

    /// Create a pivot for a session client (agents, `vgrid apply`). Same flow
    /// as the drawer's Apply — one undo step, new sheet — but synchronous and
    /// reporting the outcome instead of only showing it.
    pub(crate) fn session_create_pivot(
        &mut self,
        source: PivotSource,
        definition: PivotDefinition,
        cx: &mut Context<Self>,
    ) -> Result<String, String> {
        let table = PivotTable {
            id: self.wb(cx).next_pivot_id(),
            name: self.wb(cx).next_pivot_name(),
            source,
            definition,
            anchor_row: 0,
            anchor_col: 0,
            extent: None,
            last_refresh: None,
            stale: false,
            source_generation: None,
        };
        let id = table.id;
        let description = format!("Create {}", table.name);
        let job = PivotJob { mode: PivotPanelMode::New, activate: false, table, source_generation: 0, description };
        if !self.run_pivot_job_sync(job, cx) {
            return Err(self.status_message.clone().unwrap_or_else(|| "Could not create the pivot table.".into()));
        }
        let wb = self.wb(cx);
        let (idx, t) = wb.find_pivot(id).ok_or("The pivot table was not placed.")?;
        let (rows, cols) = t.extent.unwrap_or((0, 0));
        Ok(format!("Created {} on sheet \"{}\" (index {}): {} × {}", t.name, wb.sheets()[idx].name, idx, rows, cols))
    }

    /// Refresh one pivot for a session client. One undo step.
    pub(crate) fn session_refresh_pivot(&mut self, pivot_id: u64, cx: &mut Context<Self>) -> Result<String, String> {
        let table = self.wb(cx).find_pivot(pivot_id).map(|(_, t)| t.clone()).ok_or("This pivot table no longer exists.")?;
        let name = table.name.clone();
        let description = format!("Refresh {name}");
        let job = PivotJob { mode: PivotPanelMode::Edit { pivot_id }, activate: false, table, source_generation: 0, description };
        if !self.run_pivot_job_sync(job, cx) {
            return Err(format!("{name}: {}", self.status_message.clone().unwrap_or_else(|| "refresh failed".into())));
        }
        let (rows, cols) = self.wb(cx).find_pivot(pivot_id).and_then(|(_, t)| t.extent).unwrap_or((0, 0));
        Ok(format!("{name} ({rows} × {cols})"))
    }

    fn set_pivot_panel_message(&mut self, msg: String) {
        if let Some(p) = self.pivot_panel.as_mut() {
            p.message = Some(msg.clone());
            p.busy = false;
        }
        self.status_message = Some(msg);
    }

    /// Validate, snapshot and compute; large sources aggregate on a background
    /// thread and are placed when done.
    fn run_pivot_job(&mut self, mut job: PivotJob, cx: &mut Context<Self>) {
        if self.block_if_previewing_only(cx) { return; }
        let (table, snapshot, generation) = match self.pivot_prepare_snapshot(&job, cx) {
            Ok(s) => s,
            Err(msg) => {
                self.pivot_job_failed(&job, msg, cx);
                return;
            }
        };
        job.table = table;
        job.source_generation = generation;
        if job.table.source.data_rows() <= BACKGROUND_ROWS {
            let result = pivot::aggregate(&job.table.definition, &snapshot);
            self.pivot_finish(job, result, cx);
            return;
        }
        if let Some(p) = self.pivot_panel.as_mut() {
            p.busy = true;
            p.message = Some(format!("Computing over {} rows…", job.table.source.data_rows()));
        }
        self.status_message = Some(format!("Computing {}…", job.table.name));
        cx.notify();
        let def = job.table.definition.clone();
        cx.spawn(async move |this, cx| {
            let result = cx.background_executor().spawn(async move { pivot::aggregate(&def, &snapshot) }).await;
            let _ = this.update(cx, |this, cx| this.pivot_finish(job, result, cx));
        })
        .detach();
    }

    /// The synchronous path, for Refresh All. Returns false on failure.
    fn run_pivot_job_sync(&mut self, mut job: PivotJob, cx: &mut Context<Self>) -> bool {
        if self.block_if_previewing_only(cx) { return false; }
        match self.pivot_prepare_snapshot(&job, cx) {
            Ok((table, snapshot, generation)) => {
                job.table = table;
                job.source_generation = generation;
                let result = pivot::aggregate(&job.table.definition, &snapshot);
                self.pivot_finish(job, result, cx)
            }
            Err(msg) => {
                self.pivot_job_failed(&job, msg, cx);
                false
            }
        }
    }

    fn pivot_prepare_snapshot(&self, job: &PivotJob, cx: &App) -> Result<(PivotTable, PivotSnapshot, u64), String> {
        if job.table.definition.is_empty() {
            return Err(PivotError::NoFields.to_string());
        }
        self.wb(cx).pivot_snapshot(&job.table).map_err(|e| e.to_string())
    }

    fn pivot_job_failed(&mut self, job: &PivotJob, msg: String, cx: &mut Context<Self>) {
        if job.mode != PivotPanelMode::New {
            self.pivot_errors.insert(job.table.id, msg.clone());
        }
        self.set_pivot_panel_message(msg);
        cx.notify();
    }

    /// Place a computed output. Returns true on success. On any failure the
    /// workbook is unchanged and the last committed output stays.
    fn pivot_finish(
        &mut self,
        mut job: PivotJob,
        result: Result<PivotOutput, PivotError>,
        cx: &mut Context<Self>,
    ) -> bool {
        if self.block_if_previewing_only(cx) { return false; }
        if let Some(p) = self.pivot_panel.as_mut() {
            p.busy = false;
        }
        let output = match result {
            Ok(o) => o,
            Err(e) => {
                self.pivot_job_failed(&job, e.to_string(), cx);
                return false;
            }
        };
        // The source must not have changed while we computed.
        if self.wb(cx).pivot_source_generation(&job.table) != Some(job.source_generation) {
            self.pivot_job_failed(&job, "The source changed while computing. Refresh again.".into(), cx);
            return false;
        }
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs() as i64)
            .unwrap_or(0);
        let rows = output.height();
        let cols = output.width();

        match job.mode {
            PivotPanelMode::New => {
                format_new_pivot_values(&mut job.table.definition, &output);
                // Create the output sheet only now, so a failed or cancelled
                // create never leaves an empty sheet behind.
                let mut candidate = self.wb(cx).clone();
                let name = crate::structured_results::unique_sheet_name(&candidate, "Pivot");
                let sheet_index = candidate.add_sheet_named(&name).unwrap_or_else(|| candidate.add_sheet());
                if let Some(sheet) = candidate.sheet_mut(sheet_index) {
                    style_new_pivot(sheet, &job.table, &output);
                }
                let Some(created) = candidate.sheet(sheet_index).cloned() else { return false };
                let sheet_id = created.id;
                let created_name = created.name.clone();
                let commit = match candidate.prepare_pivot_commit(sheet_id, job.table.clone(), &output, job.source_generation, now) {
                    Ok(commit) => commit,
                    Err(error) => {
                        self.pivot_job_failed(&job, error.to_string(), cx);
                        return false;
                    }
                };
                if let Err(error) = candidate.apply_pivot_state(&commit.after) {
                    self.pivot_job_failed(&job, error.to_string(), cx);
                    return false;
                }
                self.workbook.update(cx, |wb, _| wb.restore_snapshot_monotonic(&candidate));
                if job.activate {
                    self.activate_sheet(sheet_index, cx);
                    self.row_view = RowView::new(NUM_ROWS);
                    self.clear_selection_state();
                }
                // Widths are part of creation's single undo step. Fit once;
                // refresh must preserve widths the user has subsequently set.
                let columns: Vec<usize> = (job.table.anchor_col as usize..job.table.anchor_col as usize + cols).collect();
                let widths = match self.wb(cx).sheet(sheet_index) {
                    Some(sheet) => self.measure_columns_in(sheet, &columns, None),
                    None => Default::default(),
                };
                let mut actions = vec![crate::history::UndoAction::PivotCommit {
                    commit: Box::new(commit),
                    created_sheet: Some((sheet_index, Box::new(created))),
                    description: job.description.clone(),
                }];
                for col in columns {
                    let width = widths.get(&col).copied().unwrap_or(0.0).max(self.metrics.default_cell_sizes.column_width);
                    self.set_col_width_on(sheet_id, col, width);
                    actions.push(crate::history::UndoAction::ColumnWidthSet { sheet_id, col, old: None, new: Some(width) });
                }
                self.history.record_action_with_provenance(crate::history::UndoAction::Group {
                    actions, description: job.description.clone(),
                }, None);
                if let Some(p) = self.pivot_panel.as_mut() {
                    if p.mode == PivotPanelMode::New {
                        p.mode = PivotPanelMode::Edit { pivot_id: job.table.id };
                        p.draft = job.table.definition.clone();
                    }
                    p.message = Some(format!("{} created: {rows} × {cols}.", job.table.name));
                }
                self.status_message = Some(if job.activate {
                    format!("{} created on a new sheet.", job.table.name)
                } else {
                    format!("{} created on sheet \"{}\".", job.table.name, created_name)
                });
            }
            PivotPanelMode::Edit { pivot_id } => {
                let Some((idx, _)) = self.wb(cx).find_pivot(pivot_id) else {
                    self.pivot_job_failed(&job, "This pivot table no longer exists.".into(), cx);
                    return false;
                };
                let sheet_id = self.wb(cx).sheet(idx).map(|s| s.id).unwrap();
                let commit = match self.wb(cx).prepare_pivot_commit(sheet_id, job.table.clone(), &output, job.source_generation, now) {
                    Ok(c) => c,
                    Err(e) => {
                        self.pivot_job_failed(&job, e.to_string(), cx);
                        return false;
                    }
                };
                if let Err(error) = self.workbook.update(cx, |wb, _| wb.apply_pivot_state(&commit.after)) {
                    self.pivot_job_failed(&job, error.to_string(), cx);
                    return false;
                }
                self.record_pivot_commit(commit, None, job.description.clone());
                let growth = self.wb(cx).pivot_source_growth(&job.table);
                let note = growth.map(|last| {
                    format!(
                        " {} new row{} below the source: open the field list and press X to include them.",
                        last - job.table.source.end_row,
                        if last - job.table.source.end_row == 1 { "" } else { "s" }
                    )
                });
                if let Some(p) = self.pivot_panel.as_mut() {
                    p.growth = growth;
                    p.message = Some(format!("{}: {rows} × {cols}.", job.description));
                }
                self.status_message = Some(format!("{}.{}", job.description, note.unwrap_or_default()));
            }
        }
        if let Some(old_panel) = self.pivot_panel.as_ref().filter(|p| p.mode == (PivotPanelMode::Edit { pivot_id: job.table.id })) {
            let message = old_panel.message.clone();
            if let Ok(mut panel) = self.build_pivot_panel(old_panel.mode.clone(), job.table.source, job.table.definition.clone(), cx) {
                panel.message = message;
                panel.growth = self.wb(cx).pivot_source_growth(&job.table);
                self.pivot_panel = Some(panel);
            }
        }
        self.sync_table_view(cx);
        self.pivot_errors.remove(&job.table.id);
        self.bump_cells_rev();
        self.is_modified = true;
        cx.notify();
        true
    }

    fn record_pivot_commit(
        &mut self,
        commit: PivotCommit,
        created_sheet: Option<(usize, Box<Sheet>)>,
        description: String,
    ) {
        self.history.record_action_with_provenance(
            crate::history::UndoAction::PivotCommit { commit: Box::new(commit), created_sheet, description },
            None,
        );
    }
}

// ---------------------------------------------------------------------------
// Status bar and Problems.
// ---------------------------------------------------------------------------

impl Spreadsheet {
    /// Status-bar text when the cursor is inside a pivot's output.
    pub(crate) fn pivot_status_text(&self, cx: &App) -> Option<String> {
        let t = self.pivot_under_cursor(cx)?;
        let wb = self.wb(cx);
        let source_sheet = wb.sheet_by_id(t.source.sheet_id).map(|s| s.name.clone()).unwrap_or_else(|| "?".into());
        let range = format!(
            "{}{}:{}{}",
            Spreadsheet::col_letter(t.source.start_col as usize),
            t.source.start_row + 1,
            Spreadsheet::col_letter(t.source.end_col as usize),
            t.source.end_row + 1
        );
        let state = if self.pivot_errors.contains_key(&t.id) {
            "refresh failed (see Problems)".to_string()
        } else if wb.is_pivot_stale(&t) {
            "out of date · Alt+F5 refreshes".to_string()
        } else {
            "up to date".to_string()
        };
        let source = t.source.table_id.map(|id| wb.table(id).map(|(_, t)| t.name.clone()).unwrap_or_else(|| "Missing Table".into()))
            .unwrap_or_else(|| format!("{}!{}", source_sheet, range));
        Some(format!("{} · {} · {}", t.name, source, state))
    }

    /// Pivot entries for the Problems panel: failed refreshes and stale pivots,
    /// each anchored at the pivot's top-left output cell.
    pub(crate) fn pivot_problems(&self, cx: &App) -> Vec<crate::rewind_state::Problem> {
        let wb = self.wb(cx);
        let mut out = Vec::new();
        for (sheet_idx, t) in wb.pivots() {
            let sheet_name = wb.sheet(sheet_idx).map(|s| s.name.clone()).unwrap_or_default();
            let (row, col) = (t.anchor_row as usize, t.anchor_col as usize);
            if let Some(err) = self.pivot_errors.get(&t.id) {
                out.push(crate::rewind_state::Problem {
                    sheet_idx,
                    sheet_name,
                    row,
                    col,
                    error: "Pivot refresh failed".into(),
                    formula: format!("{}: {}", t.name, err),
                });
            } else if wb.is_pivot_stale(t) {
                out.push(crate::rewind_state::Problem {
                    sheet_idx,
                    sheet_name,
                    row,
                    col,
                    error: "Pivot out of date".into(),
                    formula: format!("{}: its source changed. Alt+F5 refreshes.", t.name),
                });
            }
        }
        out
    }
}

#[cfg(test)]
mod panel_tests {
    use super::{PivotPanel, PivotPanelItem, PivotPanelMode};
    use visigrid_engine::cell::NumberFormat;
    use visigrid_engine::pivot::{Aggregation, PivotDefinition, PivotSource};
    use visigrid_engine::sheet::SheetId;

    fn panel() -> PivotPanel {
        let money = NumberFormat::Currency { decimals: 2, thousands: true, negative: Default::default(), symbol: None };
        PivotPanel {
            mode: PivotPanelMode::New,
            source: PivotSource { table_id: None, sheet_id: SheetId(1), start_row: 0, start_col: 0, end_row: 10, end_col: 2 },
            headers: vec!["Region".into(), "Customer".into(), "Amount".into()],
            column_ids: Vec::new(), table_name: None, source_menu: false, source_cursor: 0,
            column_formats: vec![NumberFormat::General, NumberFormat::General, money],
            column_numeric: vec![false, false, true],
            draft: PivotDefinition::default(),
            cursor: 0,
            message: None,
            growth: None,
            busy: false,
        }
    }

    #[::core::prelude::v1::test]
    fn assign_fields_with_sensible_defaults() {
        let mut p = panel();
        p.cursor = 0;
        assert!(p.assign('r').is_none());
        p.cursor = 2;
        assert!(p.assign('v').is_none()); // numeric → Sum, currency format
        p.cursor = 1;
        assert!(p.assign('v').is_none()); // text → Count, whole numbers
        assert_eq!(p.draft.rows[0].header, "Region");
        assert_eq!(p.draft.values[0].aggregation, Aggregation::Sum);
        assert!(matches!(p.draft.values[0].number_format, Some(NumberFormat::Currency { .. })));
        assert_eq!(p.draft.values[1].aggregation, Aggregation::Count);
        assert!(matches!(p.draft.values[1].number_format, Some(NumberFormat::Number { decimals: 0, .. })));
        // A field can't be both a row and the column field.
        p.cursor = 0;
        assert!(p.assign('c').is_some());
        // Items list: 3 fields, 1 row, 2 values.
        assert_eq!(p.items().len(), 6);
    }

    #[::core::prelude::v1::test]
    fn cycle_aggregation_moves_format_with_count_boundary_and_bulk_applies() {
        let mut p = panel();
        p.cursor = 2;
        p.assign('v');
        p.cursor = 1;
        p.assign('v');
        let first_value = p.items().iter().position(|i| *i == PivotPanelItem::Value(0)).unwrap();
        p.cursor = first_value;
        p.cycle_aggregation(true); // Sum → Count
        assert_eq!(p.draft.values[0].aggregation, Aggregation::Count);
        assert!(matches!(p.draft.values[0].number_format, Some(NumberFormat::Number { decimals: 0, .. })));
        p.cycle_aggregation(false); // back to Sum: currency again
        assert!(matches!(p.draft.values[0].number_format, Some(NumberFormat::Currency { .. })));
        // "=" applies Sum to every value field.
        assert!(p.apply_aggregation_to_all().is_some());
        assert!(p.draft.values.iter().all(|v| v.aggregation == Aggregation::Sum));
    }

    #[::core::prelude::v1::test]
    fn reorder_and_remove_keep_cursor_on_the_item() {
        let mut p = panel();
        p.cursor = 0;
        p.assign('r');
        p.cursor = 1;
        p.assign('r');
        let second_row = p.items().iter().position(|i| *i == PivotPanelItem::Row(1)).unwrap();
        p.cursor = second_row;
        p.reorder(true);
        assert_eq!(p.draft.rows[0].header, "Customer");
        assert_eq!(p.current(), Some(PivotPanelItem::Row(0)));
        p.remove_current();
        assert_eq!(p.draft.rows.len(), 1);
        assert_eq!(p.draft.rows[0].header, "Region");
    }

    #[::core::prelude::v1::test]
    fn created_style_survives_refresh_and_sheet_redo() {
        use super::{format_new_pivot_values, style_new_pivot};
        use visigrid_engine::pivot::{aggregate, PivotField, PivotTable, PivotValueField};
        use visigrid_engine::workbook::Workbook;

        let mut wb = Workbook::new();
        let source_id = wb.sheet(0).unwrap().id;
        for (r, row) in [
            ["Region", "Product", "Revenue"],
            ["East", "Desk", "125074.98"],
            ["West", "Chair", "124875.02"],
        ].iter().enumerate() {
            for (c, value) in row.iter().enumerate() {
                wb.set_cell_value_tracked(0, r, c, value);
            }
        }
        let field = |offset, header: &str| PivotField { column_id: None, offset, header: header.into() };
        let mut table = PivotTable {
            id: wb.next_pivot_id(), name: wb.next_pivot_name(),
            source: PivotSource { table_id: None, sheet_id: source_id, start_row: 0, start_col: 0, end_row: 2, end_col: 2 },
            definition: PivotDefinition {
                rows: vec![field(0, "Region")], column: Some(field(1, "Product")),
                values: vec![PivotValueField { field: field(2, "Revenue"), aggregation: Aggregation::Sum, number_format: None }],
            },
            anchor_row: 0, anchor_col: 0, extent: None, last_refresh: None,
            stale: false, source_generation: None,
        };
        let (mut table, snapshot, generation) = wb.pivot_snapshot(&table).unwrap();
        let output = aggregate(&table.definition, &snapshot).unwrap();
        format_new_pivot_values(&mut table.definition, &output);
        let index = wb.add_sheet();
        let sheet_id = wb.sheet(index).unwrap().id;
        style_new_pivot(wb.sheet_mut(index).unwrap(), &table, &output);
        let created = wb.sheet(index).unwrap().clone();
        let commit = wb.prepare_pivot_commit(sheet_id, table.clone(), &output, generation, 0).unwrap();
        wb.apply_pivot_state(&commit.after).unwrap();
        let sheet = wb.sheet(index).unwrap();
        assert!(sheet.get_format(0, 0).bold);
        assert!(sheet.get_format(1, 0).border_bottom.is_set());
        assert!(sheet.get_format(4, 0).border_top.is_set());
        assert!(sheet.get_format(2, 3).bold);
        assert!(sheet.get_format(2, 3).border_left.is_set());
        assert_eq!(sheet.get_formatted_display(2, 2), "125,074.98");
        assert_eq!(sheet.get_formatted_display(4, 3), "249,950.00");
        let formats: Vec<_> = (0..5).flat_map(|r| (0..4).map(move |c| (r, c)))
            .map(|(r, c)| (r, c, sheet.get_format(r, c))).collect();

        // Same sequence used by the desktop create action: remove the sheet,
        // restore its styled snapshot, then restore the committed values.
        wb.apply_pivot_state(&commit.before).unwrap();
        wb.take_sheet(index);
        wb.restore_sheet(index, created);
        wb.apply_pivot_state(&commit.after).unwrap();
        for (r, c, format) in formats {
            assert_eq!(wb.sheet(index).unwrap().get_format(r, c), format);
        }

        // Refresh cannot repaint a user's customized header or body cell.
        wb.sheet_mut(index).unwrap().set_bold(0, 0, false);
        wb.sheet_mut(index).unwrap().set_background_color(2, 2, Some([250, 220, 100, 255]));
        wb.set_cell_value_tracked(0, 1, 2, "125000.50");
        let (table, snapshot, generation) = wb.pivot_snapshot(&table).unwrap();
        let output = aggregate(&table.definition, &snapshot).unwrap();
        let refresh = wb.prepare_pivot_commit(sheet_id, table, &output, generation, 1).unwrap();
        wb.apply_pivot_state(&refresh.after).unwrap();
        assert!(!wb.sheet(index).unwrap().get_format(0, 0).bold);
        assert_eq!(wb.sheet(index).unwrap().get_format(2, 2).background_color, Some([250, 220, 100, 255]));
        assert_eq!(wb.sheet(index).unwrap().get_formatted_display(2, 2), "125,000.50");
    }

    #[::core::prelude::v1::test]
    fn creation_keeps_source_formats_and_does_not_mark_group_labels_as_totals() {
        use super::{format_new_pivot_values, style_new_pivot};
        use visigrid_engine::formula::eval::Value;
        use visigrid_engine::pivot::{PivotField, PivotOutput, PivotTable};
        use visigrid_engine::sheet::Sheet;

        let mut p = panel();
        p.cursor = 2;
        p.assign('v'); // currency
        p.cursor = 1;
        p.assign('v'); // count
        let original = p.draft.clone();
        let output = PivotOutput {
            cells: vec![vec![Value::Number(1234.5), Value::Number(7.0)]],
            header_rows: 0, row_groups: 0, column_items: 0, source_rows: 10,
            value_columns: vec![Some(0), Some(1)],
        };
        format_new_pivot_values(&mut p.draft, &output);
        assert_eq!(p.draft, original);

        let mut sheet = Sheet::new(SheetId(2), 100, 20);
        let table = PivotTable {
            id: 1, name: "Groups".into(), source: p.source,
            definition: PivotDefinition {
                rows: vec![PivotField { column_id: None, offset: 0, header: "Region".into() }],
                ..Default::default()
            },
            anchor_row: 3, anchor_col: 2, extent: None, last_refresh: None,
            stale: false, source_generation: None,
        };
        let output = PivotOutput {
            cells: vec![vec![Value::Text("Region".into())], vec![Value::Text("East".into())]],
            header_rows: 1, row_groups: 1, column_items: 0, source_rows: 10,
            value_columns: vec![None],
        };
        style_new_pivot(&mut sheet, &table, &output);
        assert!(sheet.get_format(3, 2).bold);
        assert!(!sheet.get_format(4, 2).bold);
        assert!(!sheet.get_format(4, 2).border_top.is_set());
        assert_eq!(sheet.get_format(0, 0), Default::default());
    }
}

#[cfg(test)]
mod table_source_tests {
    use super::*;
    use visigrid_engine::{table::{TableColumnId, TableRange}, table_view::{TableSort, TableViewSpec}, filter::SortDirection, workbook::Workbook};

    #[::core::prelude::v1::test]
    fn table_field_picker_preserves_identity_and_filtered_create_replays() {
        let mut wb = Workbook::new();
        let sheet = wb.sheet(0).unwrap().id;
        for (r, row) in [["Region", "Amount"], ["West", "10"], ["East", "20"]].iter().enumerate() {
            for (c, value) in row.iter().enumerate() { wb.set_cell_value_tracked(0, r, c, value); }
        }
        let id = wb.create_table(sheet, TableRange { start_row: 0, start_col: 0, end_row: 2, end_col: 1 }, "Sales").unwrap().table_id();
        let columns = &wb.table(id).unwrap().1.columns;
        let mut panel = PivotPanel { mode: PivotPanelMode::New, source: wb.table_pivot_source(id).unwrap(),
            headers: columns.iter().map(|c| c.name.clone()).collect(), column_ids: columns.iter().map(|c| Some(c.id)).collect(),
            table_name: Some("Sales".into()), source_menu: false, source_cursor: 0,
            column_formats: vec![NumberFormat::General; 2], column_numeric: vec![false, true],
            draft: PivotDefinition::default(), cursor: 0, message: None, growth: None, busy: false };
        panel.assign('r');
        panel.cursor = 1;
        panel.assign('v');
        assert_eq!(panel.draft.values[0].field.column_id, panel.column_ids[1]);
        let mut spec = TableViewSpec::new(id);
        spec.sort = Some(TableSort { column: columns[0].id, direction: SortDirection::Ascending });
        wb.set_table_view_spec(sheet, Some(spec.clone())).unwrap();
        let (pivot, index) = wb.create_pivot(panel.source, panel.draft).unwrap();
        let deleted = wb.prepare_pivot_delete(pivot).unwrap();
        let commit = PivotCommit { before: deleted.after, after: deleted.before };
        let mut empty = wb.clone();
        empty.apply_pivot_state(&commit.before).unwrap();
        let created = Some((index, Box::new(empty.sheet(index).unwrap().clone())));
        let undo = replay_pivot_candidate(&wb, &commit, &created, true).unwrap();
        assert_eq!(undo.sheets().len(), 1);
        assert_eq!(undo.sheet(0).unwrap().table_view_spec(), Some(&spec));
        let redo = replay_pivot_candidate(&undo, &commit, &created, false).unwrap();
        assert_eq!(redo.sheet(1).unwrap().get_raw(3, 1), "30");
        assert_eq!(redo.sheet(0).unwrap().table_view_spec(), Some(&spec));
        assert_eq!(redo.find_pivot(pivot).unwrap().1.definition.values[0].field.column_id, Some(TableColumnId(2)));
        use crate::history::UndoAction;
        let action = UndoAction::Group { actions: vec![
            UndoAction::PivotCommit { commit: Box::new(commit), created_sheet: created.clone(), description: "Create pivot".into() },
            UndoAction::ColumnWidthSet { sheet_id: created.unwrap().1.id, col: 0, old: None, new: Some(100.0) },
        ], description: "Create pivot".into() };
        assert!(is_pivot_history(&action));
        assert!(!is_pivot_history(&UndoAction::Group { actions: vec![action], description: "Arbitrary group".into() }));
    }
}
