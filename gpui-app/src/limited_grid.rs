//! Desktop-only finite view of the normal worksheet. No mutation of workbook data.
use crate::{
    app::{Spreadsheet, NUM_COLS, NUM_ROWS},
    settings::{user_settings, Setting},
};
use gpui::*;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use visigrid_engine::{
    sheet::{Sheet, SheetId},
    structural::Axis,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct GridExtent {
    pub rows: usize,
    pub cols: usize,
}
impl GridExtent {
    pub const FULL: Self = Self {
        rows: NUM_ROWS,
        cols: NUM_COLS,
    };
    const START: Self = Self { rows: 20, cols: 11 };
    fn max(self, other: Self) -> Self {
        Self {
            rows: self.rows.max(other.rows),
            cols: self.cols.max(other.cols),
        }
    }
    fn bounded(self) -> Self {
        Self {
            rows: self.rows.clamp(1, NUM_ROWS),
            cols: self.cols.clamp(1, NUM_COLS),
        }
    }
}

// Workbook identity/revision, GUI revision, sheet generation, preview mode.
type BoundsCacheKey = (EntityId, u64, u64, u64, bool);

pub(crate) struct LimitedGrid {
    pub enabled: bool,
    pub extent: GridExtent,
    cache: HashMap<SheetId, (BoundsCacheKey, GridExtent)>,
    last_row_content: Option<(SheetId, usize, bool)>,
    #[cfg(test)]
    pub scans: usize,
}
impl Default for LimitedGrid {
    fn default() -> Self {
        Self {
            enabled: false,
            extent: GridExtent::FULL,
            cache: HashMap::new(),
            last_row_content: None,
            #[cfg(test)]
            scans: 0,
        }
    }
}
impl LimitedGrid {
    pub fn invalidate(&mut self) {
        self.cache.clear();
        self.last_row_content = None;
    }
}

/// R2: sparse content and metadata only. Formatting and validation are excluded.
pub(crate) fn content_bounds(sheet: &Sheet) -> GridExtent {
    let (rows, cols) = sheet.data_bounds();
    let mut bounds = GridExtent { rows, cols };
    let mut include = |row: usize, col: usize| {
        bounds.rows = bounds.rows.max(row.saturating_add(1));
        bounds.cols = bounds.cols.max(col.saturating_add(1));
    };
    for merge in &sheet.merged_regions {
        include(merge.end.0, merge.end.1);
    }
    for table in sheet.tables() {
        let range = table.full_range();
        include(range.end_row, range.end_col);
    }
    for pivot in &sheet.pivots {
        if let Some((_, _, row, col)) = pivot.region() {
            include(row, col);
        }
    }
    for ((row, col), _) in sheet.comments() {
        include(row, col);
    }
    bounds.bounded()
}

impl Spreadsheet {
    pub(crate) fn limited_grid_enabled(&self, cx: &App) -> bool {
        self.doc_settings
            .display
            .limited_grid
            .resolve(user_settings(cx).appearance.limited_grid.resolve(false))
    }

    pub(crate) fn grid_rows(&self) -> usize {
        self.limited_grid.extent.rows
    }
    pub(crate) fn grid_cols(&self) -> usize {
        self.limited_grid.extent.cols
    }

    pub(crate) fn viewport_rows(&self) -> usize {
        if self.limited_grid.enabled {
            self.grid_rows().min(self.row_view.row_count())
        } else {
            self.row_view.row_count()
        }
    }

    fn cached_content_bounds(&mut self, sheet_index: usize, cx: &App) -> GridExtent {
        let sheet = self.display_workbook(cx).sheet(sheet_index).unwrap();
        let id = sheet.id;
        let key = (
            self.workbook.entity_id(),
            self.display_workbook(cx).revision(),
            self.cells_rev,
            sheet.edit_generation(),
            self.is_previewing(),
        );
        if let Some((cached, bounds)) = self.limited_grid.cache.get(&id) {
            if *cached == key {
                return *bounds;
            }
        }
        let mut bounds = content_bounds(sheet);
        // A saved sort can move canonical content to a later view slot. Include
        // its projection, cached with the content rather than walking it per frame.
        if sheet_index == self.sheet_index(cx)
            && (self.filter_state.sort.is_some() || self.table_view_installed)
        {
            bounds.rows = bounds.rows.max(
                (0..bounds.rows.min(self.row_view.row_count()))
                    .map(|row| self.row_view.data_to_view_unchecked(row) + 1)
                    .max()
                    .unwrap_or(1),
            );
        }
        self.limited_grid.cache.insert(id, (key, bounds));
        self.limited_grid.last_row_content = None;
        #[cfg(test)]
        {
            self.limited_grid.scans += 1;
        }
        bounds
    }

    pub(crate) fn sheet_grid_extent(&mut self, index: usize, cx: &App) -> GridExtent {
        if !self.limited_grid_enabled(cx) {
            return GridExtent::FULL;
        }
        let content = self.cached_content_bounds(index, cx);
        let sheet = self.display_workbook(cx).sheet(index).unwrap();
        let id = sheet.id.0;
        let frozen = GridExtent {
            rows: sheet.frozen_panes.0.saturating_add(1),
            cols: sheet.frozen_panes.1.saturating_add(1),
        };
        let stored = self
            .doc_settings
            .display
            .sheet_extents
            .entry(id)
            .or_insert(content.max(GridExtent::START));
        *stored = stored.max(frozen).bounded();
        stored.max(content).bounded()
    }

    /// Constant-time cache lookup when unchanged. Called before view/navigation work.
    pub(crate) fn sync_limited_grid(&mut self, cx: &App) {
        if !self.doc_settings.display.sheet_extents.is_empty() {
            let ids: std::collections::HashSet<_> = self
                .wb(cx)
                .sheets()
                .iter()
                .chain(self.display_workbook(cx).sheets())
                .map(|s| s.id.0)
                .collect();
            self.doc_settings
                .display
                .sheet_extents
                .retain(|id, _| ids.contains(id));
        }
        self.limited_grid.enabled = self.limited_grid_enabled(cx);
        self.limited_grid.extent = self.sheet_grid_extent(self.sheet_index(cx), cx);
        if !self.limited_grid.enabled {
            return;
        }
        let extent = self.limited_grid.extent;
        let editing = self.mode.is_editing();
        let clamp = |view: &mut crate::workbook_view::WorkbookViewState| {
            if !editing {
                view.selected.0 = view.selected.0.min(extent.rows - 1);
                view.selected.1 = view.selected.1.min(extent.cols - 1);
            }
            view.scroll_row = view
                .scroll_row
                .min(extent.rows - 1)
                .max(view.frozen_rows.min(extent.rows - 1));
            view.scroll_col = view
                .scroll_col
                .min(extent.cols - 1)
                .max(view.frozen_cols.min(extent.cols - 1));
            // Full-row/column/sheet selections retain their data-model extent.
        };
        clamp(&mut self.view_state);
        if let Some(pane) = &mut self.split_pane {
            clamp(&mut pane.view_state);
        }
    }

    pub(crate) fn toggle_limited_grid(&mut self, cx: &mut Context<Self>) {
        self.commit_pending_edit(cx);
        if self.mode.is_editing() {
            return;
        }
        self.doc_settings.display.limited_grid = Setting::Value(!self.limited_grid_enabled(cx));
        self.sync_limited_grid(cx);
        self.save_doc_settings_if_needed();
        cx.notify();
    }

    pub(crate) fn set_grid_extent(&mut self, extent: GridExtent, cx: &mut Context<Self>) {
        let content = self.cached_content_bounds(self.sheet_index(cx), cx);
        let sheet = self.sheet(cx);
        let id = sheet.id.0;
        let floor = GridExtent {
            rows: sheet.frozen_panes.0 + 1,
            cols: sheet.frozen_panes.1 + 1,
        };
        self.doc_settings
            .display
            .sheet_extents
            .insert(id, extent.max(content).max(floor).bounded());
        self.sync_limited_grid(cx);
        self.save_doc_settings_if_needed();
        cx.notify();
    }

    /// Structural edits change an existing view extent even while the view is off.
    pub(crate) fn adjust_grid_structure(
        &mut self,
        index: usize,
        axis: Axis,
        count: usize,
        delete: bool,
        cx: &mut Context<Self>,
    ) {
        let Some(sheet) = self.wb(cx).sheet(index) else {
            return;
        };
        let id = sheet.id.0;
        let Some(mut extent) = self.doc_settings.display.sheet_extents.get(&id).copied() else {
            return;
        };
        let value = match axis {
            Axis::Row => &mut extent.rows,
            Axis::Col => &mut extent.cols,
        };
        *value = if delete {
            value.saturating_sub(count)
        } else {
            value.saturating_add(count)
        };
        let content = self.cached_content_bounds(index, cx);
        let sheet = self.wb(cx).sheet(index).unwrap();
        let frozen = GridExtent {
            rows: sheet.frozen_panes.0 + 1,
            cols: sheet.frozen_panes.1 + 1,
        };
        self.doc_settings
            .display
            .sheet_extents
            .insert(id, extent.max(content).max(frozen).bounded());
        self.sync_limited_grid(cx);
        self.save_doc_settings_if_needed();
    }

    pub(crate) fn grow_grid(&mut self, rows: usize, cols: usize, cx: &mut Context<Self>) {
        self.sync_limited_grid(cx);
        if !self.limited_grid.enabled {
            return;
        }
        self.set_grid_extent(
            GridExtent {
                rows: self.grid_rows().saturating_add(rows),
                cols: self.grid_cols().saturating_add(cols),
            },
            cx,
        );
    }

    pub(crate) fn shrink_grid_to_data(&mut self, cx: &mut Context<Self>) {
        self.commit_pending_edit(cx);
        if self.mode.is_editing() {
            return;
        }
        let content = self.cached_content_bounds(self.sheet_index(cx), cx);
        self.set_grid_extent(content, cx);
    }

    /// Only a one-row downward move can extend a nonempty last row (R4).
    pub(crate) fn grow_grid_for_down(&mut self, cx: &mut Context<Self>) {
        self.sync_limited_grid(cx);
        let row = self.active_view_state().selected.0;
        if !self.limited_grid.enabled || row + 1 != self.grid_rows() || row + 1 >= NUM_ROWS {
            return;
        }
        let data_row = self.row_view.view_to_data(row);
        // A row has content if its own R2 metadata/data intersects it.
        let id = self.sheet(cx).id;
        let has_content = match self.limited_grid.last_row_content {
            Some((cached_id, cached_row, value)) if cached_id == id && cached_row == data_row => {
                value
            }
            _ => {
                let value = row_has_content(self.sheet(cx), data_row);
                self.limited_grid.last_row_content = Some((id, data_row, value));
                value
            }
        };
        if has_content {
            self.grow_grid(1, 0, cx);
        }
    }

    pub(crate) fn allow_grid_jump(
        &mut self,
        sheet: usize,
        row: usize,
        col: usize,
        cx: &mut Context<Self>,
    ) -> bool {
        let extent = self.sheet_grid_extent(sheet, cx);
        if row < extent.rows && col < extent.cols {
            return true;
        }
        self.status_message = Some(format!("{} is outside this sheet's grid. Add rows or columns to reach it, or turn off Limited Grid.", self.cell_ref_at(row, col)));
        cx.notify();
        false
    }

    pub(crate) fn show_add_count(&mut self, axis: Axis, cx: &mut Context<Self>) {
        self.sync_limited_grid(cx);
        let (title, max) = match axis {
            Axis::Row => ("Add rows", NUM_ROWS - self.grid_rows()),
            Axis::Col => ("Add columns", NUM_COLS - self.grid_cols()),
        };
        if max == 0 {
            self.status_message = Some("This grid is already at the worksheet limit.".into());
            cx.notify();
            return;
        }
        self.count_dialog = Some(crate::count_dialog::CountDialog {
            title,
            axis,
            at: 0,
            max,
            input: "1".into(),
            input_selected: true,
            grow_only: true,
            anchor: None,
        });
        self.mode = crate::mode::Mode::CountDialog;
        cx.notify();
    }
}

fn row_has_content(sheet: &Sheet, row: usize) -> bool {
    // This runs only when attempting to leave the last row, never during rendering.
    sheet
        .cells_iter()
        .any(|((r, _), cell)| r == row && (!cell.value().is_empty() || cell.comment().is_some()))
        || sheet.spill_receiver_coords().any(|(r, _)| r == row)
        || sheet
            .merged_regions
            .iter()
            .any(|m| m.start.0 <= row && row <= m.end.0)
        || sheet.tables().iter().any(|t| {
            let r = t.full_range();
            r.start_row <= row && row <= r.end_row
        })
        || sheet.pivots.iter().any(|p| {
            p.region()
                .is_some_and(|(r, _, end, _)| r <= row && row <= end)
        })
}
