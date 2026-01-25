//! Undo/Redo operations for Spreadsheet.
//!
//! This module contains:
//! - undo() and redo() public methods
//! - apply_undo_action() and apply_redo_action() helper methods
//! - Handles all UndoAction variants: Values, Format, NamedRange*, Group, Rows*, Cols*, Sort

use gpui::*;

use crate::app::Spreadsheet;
use crate::history::UndoAction;

/// Maximum rows in the spreadsheet
const NUM_ROWS: usize = 1_000_000;
/// Maximum columns in the spreadsheet
const NUM_COLS: usize = 16_384;

impl Spreadsheet {
    // Undo/Redo
    pub fn undo(&mut self, cx: &mut Context<Self>) {
        if let Some(entry) = self.history.undo() {
            match entry.action {
                UndoAction::Values { sheet_index, changes } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for change in changes {
                            sheet.set_value(change.row, change.col, &change.old_value);
                        }
                    }
                    self.bump_cells_rev();  // Invalidate cell search cache
                    self.status_message = Some("Undo".to_string());
                }
                UndoAction::Format { sheet_index, patches, description, .. } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for patch in patches {
                            sheet.set_format(patch.row, patch.col, patch.before);
                        }
                    }
                    self.status_message = Some(format!("Undo: {}", description));
                }
                UndoAction::NamedRangeDeleted { named_range } => {
                    // Restore the deleted named range
                    let name = named_range.name.clone();
                    let _ = self.workbook.named_ranges_mut().set(named_range);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: restored '{}'", name));
                }
                UndoAction::NamedRangeCreated { name } => {
                    // Delete the created named range
                    self.workbook.delete_named_range(&name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: removed '{}'", name));
                }
                UndoAction::NamedRangeRenamed { old_name, new_name } => {
                    // Rename back to original name
                    let _ = self.workbook.rename_named_range(&new_name, &old_name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: renamed back to '{}'", old_name));
                }
                UndoAction::NamedRangeDescriptionChanged { name, old_description, .. } => {
                    // Restore the old description
                    let _ = self.workbook.named_ranges_mut().set_description(&name, old_description.clone());
                    self.status_message = Some(format!("Undo: description of '{}'", name));
                }
                UndoAction::Group { actions, description } => {
                    // Undo all actions in reverse order
                    for action in actions.into_iter().rev() {
                        self.apply_undo_action(action);
                    }
                    self.status_message = Some(format!("Undo: {}", description));
                }
                UndoAction::RowsInserted { sheet_index, at_row, count } => {
                    // Undo insert by deleting the rows
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.delete_rows(at_row, count);
                    }
                    // Shift row heights back up
                    let heights_to_shift: Vec<_> = self.row_heights
                        .iter()
                        .filter(|(r, _)| **r >= at_row + count)
                        .map(|(r, h)| (*r, *h))
                        .collect();
                    for r in at_row..NUM_ROWS {
                        self.row_heights.remove(&r);
                    }
                    for (r, h) in heights_to_shift {
                        self.row_heights.insert(r - count, h);
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: inserted {} row(s)", count));
                }
                UndoAction::RowsDeleted { sheet_index, at_row, count, deleted_cells, deleted_row_heights } => {
                    // Undo delete by re-inserting rows and restoring data
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.insert_rows(at_row, count);
                        // Restore the deleted cells
                        for (row, col, value, format) in deleted_cells {
                            sheet.set_value(row, col, &value);
                            sheet.set_format(row, col, format);
                        }
                    }
                    // Shift row heights down and restore deleted heights
                    let heights_to_shift: Vec<_> = self.row_heights
                        .iter()
                        .filter(|(r, _)| **r >= at_row)
                        .map(|(r, h)| (*r, *h))
                        .collect();
                    for (r, _) in &heights_to_shift {
                        self.row_heights.remove(r);
                    }
                    for (r, h) in heights_to_shift {
                        if r + count < NUM_ROWS {
                            self.row_heights.insert(r + count, h);
                        }
                    }
                    // Restore deleted row heights
                    for (r, h) in deleted_row_heights {
                        self.row_heights.insert(r, h);
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: deleted {} row(s)", count));
                }
                UndoAction::ColsInserted { sheet_index, at_col, count } => {
                    // Undo insert by deleting the columns
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.delete_cols(at_col, count);
                    }
                    // Shift column widths back left
                    let widths_to_shift: Vec<_> = self.col_widths
                        .iter()
                        .filter(|(c, _)| **c >= at_col + count)
                        .map(|(c, w)| (*c, *w))
                        .collect();
                    for c in at_col..NUM_COLS {
                        self.col_widths.remove(&c);
                    }
                    for (c, w) in widths_to_shift {
                        self.col_widths.insert(c - count, w);
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: inserted {} column(s)", count));
                }
                UndoAction::ColsDeleted { sheet_index, at_col, count, deleted_cells, deleted_col_widths } => {
                    // Undo delete by re-inserting columns and restoring data
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.insert_cols(at_col, count);
                        // Restore the deleted cells
                        for (row, col, value, format) in deleted_cells {
                            sheet.set_value(row, col, &value);
                            sheet.set_format(row, col, format);
                        }
                    }
                    // Shift column widths right and restore deleted widths
                    let widths_to_shift: Vec<_> = self.col_widths
                        .iter()
                        .filter(|(c, _)| **c >= at_col)
                        .map(|(c, w)| (*c, *w))
                        .collect();
                    for (c, _) in &widths_to_shift {
                        self.col_widths.remove(c);
                    }
                    for (c, w) in widths_to_shift {
                        if c + count < NUM_COLS {
                            self.col_widths.insert(c + count, w);
                        }
                    }
                    // Restore deleted column widths
                    for (c, w) in deleted_col_widths {
                        self.col_widths.insert(c, w);
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Undo: deleted {} column(s)", count));
                }
                UndoAction::SortApplied { previous_row_order, previous_sort_state, .. } => {
                    // Restore previous row order
                    self.row_view.apply_sort(previous_row_order);
                    // Restore previous sort state
                    self.filter_state.sort = previous_sort_state.map(|(col, is_ascending)| {
                        visigrid_engine::filter::SortState {
                            column: col,
                            direction: if is_ascending {
                                visigrid_engine::filter::SortDirection::Ascending
                            } else {
                                visigrid_engine::filter::SortDirection::Descending
                            },
                        }
                    });
                    self.filter_state.invalidate_all_caches();
                    // Clamp selection to valid bounds after row order change
                    self.clamp_selection();
                    self.status_message = Some("Undo: sort".to_string());
                }
            }
            self.is_modified = true;
            self.request_title_refresh(cx);
        }
    }

    /// Apply a single undo action (helper for Group handling)
    fn apply_undo_action(&mut self, action: UndoAction) {
        match action {
            UndoAction::Values { sheet_index, changes } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    // CRITICAL: Apply in reverse order to handle same-cell sequences correctly.
                    // If cell X was changed A→B→C in one batch, we must undo C→B first, then B→A.
                    for change in changes.iter().rev() {
                        sheet.set_value(change.row, change.col, &change.old_value);
                    }
                }
                self.bump_cells_rev();
            }
            UndoAction::Format { sheet_index, patches, .. } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    for patch in patches {
                        sheet.set_format(patch.row, patch.col, patch.before);
                    }
                }
            }
            UndoAction::NamedRangeDeleted { named_range } => {
                let _ = self.workbook.named_ranges_mut().set(named_range);
                self.bump_cells_rev();
            }
            UndoAction::NamedRangeCreated { name } => {
                self.workbook.delete_named_range(&name);
                self.bump_cells_rev();
            }
            UndoAction::NamedRangeRenamed { old_name, new_name } => {
                let _ = self.workbook.rename_named_range(&new_name, &old_name);
                self.bump_cells_rev();
            }
            UndoAction::NamedRangeDescriptionChanged { name, old_description, .. } => {
                let _ = self.workbook.named_ranges_mut().set_description(&name, old_description.clone());
            }
            UndoAction::Group { actions, .. } => {
                // Recursively undo nested groups
                for sub_action in actions.into_iter().rev() {
                    self.apply_undo_action(sub_action);
                }
            }
            UndoAction::RowsInserted { sheet_index, at_row, count } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.delete_rows(at_row, count);
                }
                // Shift row heights back up
                let heights_to_shift: Vec<_> = self.row_heights
                    .iter()
                    .filter(|(r, _)| **r >= at_row + count)
                    .map(|(r, h)| (*r, *h))
                    .collect();
                for r in at_row..NUM_ROWS {
                    self.row_heights.remove(&r);
                }
                for (r, h) in heights_to_shift {
                    self.row_heights.insert(r - count, h);
                }
                self.bump_cells_rev();
            }
            UndoAction::RowsDeleted { sheet_index, at_row, count, deleted_cells, deleted_row_heights } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.insert_rows(at_row, count);
                    for (row, col, value, format) in deleted_cells {
                        sheet.set_value(row, col, &value);
                        sheet.set_format(row, col, format);
                    }
                }
                // Shift row heights down and restore deleted heights
                let heights_to_shift: Vec<_> = self.row_heights
                    .iter()
                    .filter(|(r, _)| **r >= at_row)
                    .map(|(r, h)| (*r, *h))
                    .collect();
                for (r, _) in &heights_to_shift {
                    self.row_heights.remove(r);
                }
                for (r, h) in heights_to_shift {
                    if r + count < NUM_ROWS {
                        self.row_heights.insert(r + count, h);
                    }
                }
                for (r, h) in deleted_row_heights {
                    self.row_heights.insert(r, h);
                }
                self.bump_cells_rev();
            }
            UndoAction::ColsInserted { sheet_index, at_col, count } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.delete_cols(at_col, count);
                }
                // Shift column widths back left
                let widths_to_shift: Vec<_> = self.col_widths
                    .iter()
                    .filter(|(c, _)| **c >= at_col + count)
                    .map(|(c, w)| (*c, *w))
                    .collect();
                for c in at_col..NUM_COLS {
                    self.col_widths.remove(&c);
                }
                for (c, w) in widths_to_shift {
                    self.col_widths.insert(c - count, w);
                }
                self.bump_cells_rev();
            }
            UndoAction::ColsDeleted { sheet_index, at_col, count, deleted_cells, deleted_col_widths } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.insert_cols(at_col, count);
                    for (row, col, value, format) in deleted_cells {
                        sheet.set_value(row, col, &value);
                        sheet.set_format(row, col, format);
                    }
                }
                // Shift column widths right and restore deleted widths
                let widths_to_shift: Vec<_> = self.col_widths
                    .iter()
                    .filter(|(c, _)| **c >= at_col)
                    .map(|(c, w)| (*c, *w))
                    .collect();
                for (c, _) in &widths_to_shift {
                    self.col_widths.remove(c);
                }
                for (c, w) in widths_to_shift {
                    if c + count < NUM_COLS {
                        self.col_widths.insert(c + count, w);
                    }
                }
                for (c, w) in deleted_col_widths {
                    self.col_widths.insert(c, w);
                }
                self.bump_cells_rev();
            }
            UndoAction::SortApplied { previous_row_order, previous_sort_state, .. } => {
                self.row_view.apply_sort(previous_row_order);
                self.filter_state.sort = previous_sort_state.map(|(col, is_ascending)| {
                    visigrid_engine::filter::SortState {
                        column: col,
                        direction: if is_ascending {
                            visigrid_engine::filter::SortDirection::Ascending
                        } else {
                            visigrid_engine::filter::SortDirection::Descending
                        },
                    }
                });
                self.filter_state.invalidate_all_caches();
                self.clamp_selection();
            }
        }
    }

    /// Apply a single redo action (helper for Group handling)
    fn apply_redo_action(&mut self, action: UndoAction) {
        match action {
            UndoAction::Values { sheet_index, changes } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    for change in changes {
                        sheet.set_value(change.row, change.col, &change.new_value);
                    }
                }
                self.bump_cells_rev();
            }
            UndoAction::Format { sheet_index, patches, .. } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    for patch in patches {
                        sheet.set_format(patch.row, patch.col, patch.after);
                    }
                }
            }
            UndoAction::NamedRangeDeleted { named_range } => {
                let name = named_range.name.clone();
                self.workbook.delete_named_range(&name);
                self.bump_cells_rev();
            }
            UndoAction::NamedRangeCreated { .. } => {
                // Re-create is not possible without original data
            }
            UndoAction::NamedRangeRenamed { old_name, new_name } => {
                let _ = self.workbook.rename_named_range(&old_name, &new_name);
                self.bump_cells_rev();
            }
            UndoAction::NamedRangeDescriptionChanged { name, new_description, .. } => {
                let _ = self.workbook.named_ranges_mut().set_description(&name, new_description.clone());
            }
            UndoAction::Group { actions, .. } => {
                // Recursively redo nested groups
                for sub_action in actions {
                    self.apply_redo_action(sub_action);
                }
            }
            UndoAction::RowsInserted { sheet_index, at_row, count } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.insert_rows(at_row, count);
                }
                // Shift row heights down (same as insert_rows in main code)
                let heights_to_shift: Vec<_> = self.row_heights
                    .iter()
                    .filter(|(r, _)| **r >= at_row)
                    .map(|(r, h)| (*r, *h))
                    .collect();
                for (r, _) in &heights_to_shift {
                    self.row_heights.remove(r);
                }
                for (r, h) in heights_to_shift {
                    if r + count < NUM_ROWS {
                        self.row_heights.insert(r + count, h);
                    }
                }
                self.bump_cells_rev();
            }
            UndoAction::RowsDeleted { sheet_index, at_row, count, .. } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.delete_rows(at_row, count);
                }
                // Shift row heights up (same as delete_rows in main code)
                let heights_to_shift: Vec<_> = self.row_heights
                    .iter()
                    .filter(|(r, _)| **r >= at_row + count)
                    .map(|(r, h)| (*r, *h))
                    .collect();
                for r in at_row..NUM_ROWS {
                    self.row_heights.remove(&r);
                }
                for (r, h) in heights_to_shift {
                    self.row_heights.insert(r - count, h);
                }
                self.bump_cells_rev();
            }
            UndoAction::ColsInserted { sheet_index, at_col, count } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.insert_cols(at_col, count);
                }
                // Shift column widths right (same as insert_cols in main code)
                let widths_to_shift: Vec<_> = self.col_widths
                    .iter()
                    .filter(|(c, _)| **c >= at_col)
                    .map(|(c, w)| (*c, *w))
                    .collect();
                for (c, _) in &widths_to_shift {
                    self.col_widths.remove(c);
                }
                for (c, w) in widths_to_shift {
                    if c + count < NUM_COLS {
                        self.col_widths.insert(c + count, w);
                    }
                }
                self.bump_cells_rev();
            }
            UndoAction::ColsDeleted { sheet_index, at_col, count, .. } => {
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.delete_cols(at_col, count);
                }
                // Shift column widths left (same as delete_cols in main code)
                let widths_to_shift: Vec<_> = self.col_widths
                    .iter()
                    .filter(|(c, _)| **c >= at_col + count)
                    .map(|(c, w)| (*c, *w))
                    .collect();
                for c in at_col..NUM_COLS {
                    self.col_widths.remove(&c);
                }
                for (c, w) in widths_to_shift {
                    self.col_widths.insert(c - count, w);
                }
                self.bump_cells_rev();
            }
            UndoAction::SortApplied { new_row_order, new_sort_state, .. } => {
                self.row_view.apply_sort(new_row_order);
                let (col, is_ascending) = new_sort_state;
                self.filter_state.sort = Some(visigrid_engine::filter::SortState {
                    column: col,
                    direction: if is_ascending {
                        visigrid_engine::filter::SortDirection::Ascending
                    } else {
                        visigrid_engine::filter::SortDirection::Descending
                    },
                });
                self.filter_state.invalidate_all_caches();
                self.clamp_selection();
            }
        }
    }

    pub fn redo(&mut self, cx: &mut Context<Self>) {
        if let Some(entry) = self.history.redo() {
            match entry.action {
                UndoAction::Values { sheet_index, changes } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for change in changes {
                            sheet.set_value(change.row, change.col, &change.new_value);
                        }
                    }
                    self.bump_cells_rev();  // Invalidate cell search cache
                    self.status_message = Some("Redo".to_string());
                }
                UndoAction::Format { sheet_index, patches, description, .. } => {
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for patch in patches {
                            sheet.set_format(patch.row, patch.col, patch.after);
                        }
                    }
                    self.status_message = Some(format!("Redo: {}", description));
                }
                UndoAction::NamedRangeDeleted { named_range } => {
                    // Re-delete the named range
                    let name = named_range.name.clone();
                    self.workbook.delete_named_range(&name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: deleted '{}'", name));
                }
                UndoAction::NamedRangeCreated { ref name } => {
                    // Re-create is not possible without the original data
                    // This shouldn't happen in practice (create followed by undo-redo)
                    self.status_message = Some(format!("Redo: recreate '{}' not supported", name));
                }
                UndoAction::NamedRangeRenamed { old_name, new_name } => {
                    // Rename again to new name
                    let _ = self.workbook.rename_named_range(&old_name, &new_name);
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: renamed to '{}'", new_name));
                }
                UndoAction::NamedRangeDescriptionChanged { name, new_description, .. } => {
                    // Apply the new description
                    let _ = self.workbook.named_ranges_mut().set_description(&name, new_description.clone());
                    self.status_message = Some(format!("Redo: description of '{}'", name));
                }
                UndoAction::Group { actions, description } => {
                    // Redo all actions in order
                    for action in actions {
                        self.apply_redo_action(action);
                    }
                    self.status_message = Some(format!("Redo: {}", description));
                }
                UndoAction::RowsInserted { sheet_index, at_row, count } => {
                    // Re-insert the rows
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.insert_rows(at_row, count);
                    }
                    // Shift row heights down
                    let heights_to_shift: Vec<_> = self.row_heights
                        .iter()
                        .filter(|(r, _)| **r >= at_row)
                        .map(|(r, h)| (*r, *h))
                        .collect();
                    for (r, _) in &heights_to_shift {
                        self.row_heights.remove(r);
                    }
                    for (r, h) in heights_to_shift {
                        if r + count < NUM_ROWS {
                            self.row_heights.insert(r + count, h);
                        }
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: insert {} row(s)", count));
                }
                UndoAction::RowsDeleted { sheet_index, at_row, count, .. } => {
                    // Re-delete the rows
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.delete_rows(at_row, count);
                    }
                    // Shift row heights up
                    let heights_to_shift: Vec<_> = self.row_heights
                        .iter()
                        .filter(|(r, _)| **r >= at_row + count)
                        .map(|(r, h)| (*r, *h))
                        .collect();
                    for r in at_row..NUM_ROWS {
                        self.row_heights.remove(&r);
                    }
                    for (r, h) in heights_to_shift {
                        self.row_heights.insert(r - count, h);
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: delete {} row(s)", count));
                }
                UndoAction::ColsInserted { sheet_index, at_col, count } => {
                    // Re-insert the columns
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.insert_cols(at_col, count);
                    }
                    // Shift column widths right
                    let widths_to_shift: Vec<_> = self.col_widths
                        .iter()
                        .filter(|(c, _)| **c >= at_col)
                        .map(|(c, w)| (*c, *w))
                        .collect();
                    for (c, _) in &widths_to_shift {
                        self.col_widths.remove(c);
                    }
                    for (c, w) in widths_to_shift {
                        if c + count < NUM_COLS {
                            self.col_widths.insert(c + count, w);
                        }
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: insert {} column(s)", count));
                }
                UndoAction::ColsDeleted { sheet_index, at_col, count, .. } => {
                    // Re-delete the columns
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.delete_cols(at_col, count);
                    }
                    // Shift column widths left
                    let widths_to_shift: Vec<_> = self.col_widths
                        .iter()
                        .filter(|(c, _)| **c >= at_col + count)
                        .map(|(c, w)| (*c, *w))
                        .collect();
                    for c in at_col..NUM_COLS {
                        self.col_widths.remove(&c);
                    }
                    for (c, w) in widths_to_shift {
                        self.col_widths.insert(c - count, w);
                    }
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: delete {} column(s)", count));
                }
                UndoAction::SortApplied { new_row_order, new_sort_state, .. } => {
                    // Re-apply the sort
                    self.row_view.apply_sort(new_row_order);
                    let (col, is_ascending) = new_sort_state;
                    self.filter_state.sort = Some(visigrid_engine::filter::SortState {
                        column: col,
                        direction: if is_ascending {
                            visigrid_engine::filter::SortDirection::Ascending
                        } else {
                            visigrid_engine::filter::SortDirection::Descending
                        },
                    });
                    self.filter_state.invalidate_all_caches();
                    // Clamp selection to valid bounds after row order change
                    self.clamp_selection();
                    self.status_message = Some("Redo: sort".to_string());
                }
            }
            self.is_modified = true;
            self.request_title_refresh(cx);
        }
    }
}
