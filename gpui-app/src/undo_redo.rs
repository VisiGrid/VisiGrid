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
                UndoAction::NamedRangeCreated { named_range } => {
                    // Delete the created named range
                    let name = &named_range.name;
                    self.workbook.delete_named_range(name);
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
                UndoAction::SortCleared { previous_row_order, previous_sort_state, .. } => {
                    // Undo: restore previous sort state
                    self.row_view.apply_sort(previous_row_order);
                    let (col, is_ascending) = previous_sort_state;
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
                    self.status_message = Some("Undo: clear sort".to_string());
                }
                UndoAction::ValidationSet { sheet_index, range, previous_rules, .. } => {
                    // Undo: clear the new rule, restore previous rules
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        // Clear the rule that was set
                        sheet.validations.clear_range(&range);
                        // Restore previous rules that were overwritten
                        for (rule_range, rule) in previous_rules {
                            sheet.validations.set(rule_range, rule);
                        }
                    }
                    self.bump_cells_rev();
                    // Recompute invalid markers for affected range
                    self.revalidate_range(&range);
                    self.status_message = Some("Undo: set validation".to_string());
                }
                UndoAction::ValidationCleared { sheet_index, range, cleared_rules } => {
                    // Undo: restore the cleared rules
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for (rule_range, rule) in cleared_rules {
                            sheet.validations.set(rule_range, rule);
                        }
                    }
                    self.bump_cells_rev();
                    // Recompute invalid markers for affected range
                    self.revalidate_range(&range);
                    self.status_message = Some(format!("Undo: clear validation ({} cells)", range.cell_count()));
                }
                UndoAction::ValidationExcluded { sheet_index, range } => {
                    // Undo: remove the exclusion
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.validations.remove_exclusion(&range);
                    }
                    self.bump_cells_rev();
                    // Recompute invalid markers (cells may now be validated again)
                    self.revalidate_range(&range);
                    self.status_message = Some(format!("Undo: exclude from validation ({} cells)", range.cell_count()));
                }
                UndoAction::ValidationExclusionCleared { sheet_index, range, cleared_exclusions } => {
                    // Undo: restore the cleared exclusions
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        for exclusion_range in cleared_exclusions {
                            sheet.validations.exclude(exclusion_range);
                        }
                    }
                    self.bump_cells_rev();
                    // Recompute invalid markers (cells may be excluded again)
                    self.revalidate_range(&range);
                    self.status_message = Some(format!("Undo: clear exclusions ({} cells)", range.cell_count()));
                }
                UndoAction::Rewind { .. } => {
                    // Rewind is audit-only - cannot be undone
                    // (It's the last action after truncation, so this shouldn't be reached)
                    self.status_message = Some("Cannot undo rewind".to_string());
                    return;  // Don't modify state
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
            UndoAction::NamedRangeCreated { named_range } => {
                self.workbook.delete_named_range(&named_range.name);
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
            UndoAction::SortCleared { previous_row_order, previous_sort_state, .. } => {
                // Undo: restore previous sort state
                self.row_view.apply_sort(previous_row_order);
                let (col, is_ascending) = previous_sort_state;
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
            UndoAction::ValidationSet { sheet_index, range, previous_rules, .. } => {
                // Undo: clear the new rule, restore previous rules
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.validations.clear_range(&range);
                    for (rule_range, rule) in previous_rules {
                        sheet.validations.set(rule_range, rule);
                    }
                }
                self.bump_cells_rev();
                self.revalidate_range(&range);
            }
            UndoAction::ValidationCleared { sheet_index, range, cleared_rules } => {
                // Undo: restore the cleared rules
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    for (rule_range, rule) in cleared_rules {
                        sheet.validations.set(rule_range, rule);
                    }
                }
                self.bump_cells_rev();
                self.revalidate_range(&range);
            }
            UndoAction::ValidationExcluded { sheet_index, range } => {
                // Undo: remove the exclusion
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.validations.remove_exclusion(&range);
                }
                self.bump_cells_rev();
                self.revalidate_range(&range);
            }
            UndoAction::ValidationExclusionCleared { sheet_index, range, cleared_exclusions } => {
                // Undo: restore the cleared exclusions
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    for exclusion_range in cleared_exclusions {
                        sheet.validations.exclude(exclusion_range);
                    }
                }
                self.bump_cells_rev();
                self.revalidate_range(&range);
            }
            UndoAction::Rewind { .. } => {
                // Rewind is audit-only - cannot be undone
                // This should never be reached (Rewind is always last in stack)
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
            UndoAction::SortCleared { .. } => {
                // Redo: clear the sort again
                self.row_view.clear_sort();
                self.filter_state.sort = None;
                self.filter_state.invalidate_all_caches();
                self.clamp_selection();
            }
            UndoAction::ValidationSet { sheet_index, range, new_rule, .. } => {
                // Redo: full replace pipeline - clear overlaps then set rule
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.validations.clear_range(&range);
                    sheet.validations.set(range, new_rule);
                }
                self.bump_cells_rev();
                self.revalidate_range(&range);
            }
            UndoAction::ValidationCleared { sheet_index, range, .. } => {
                // Redo: clear the validations again
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.validations.clear_range(&range);
                }
                self.bump_cells_rev();
                self.clear_invalid_markers_in_range(&range);
            }
            UndoAction::ValidationExcluded { sheet_index, range } => {
                // Redo: re-add the exclusion
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.validations.exclude(range);
                }
                self.bump_cells_rev();
                self.clear_invalid_markers_in_range(&range);
            }
            UndoAction::ValidationExclusionCleared { sheet_index, range, .. } => {
                // Redo: clear exclusions again
                if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                    sheet.validations.clear_exclusions_in_range(&range);
                }
                self.bump_cells_rev();
                self.revalidate_range(&range);
            }
            UndoAction::Rewind { .. } => {
                // Rewind is audit-only - cannot be redone
                // This should never be reached
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
                UndoAction::NamedRangeCreated { ref named_range } => {
                    // Re-create the named range (now possible with full payload)
                    let name = named_range.name.clone();
                    let _ = self.workbook.named_ranges_mut().set(named_range.clone());
                    self.bump_cells_rev();
                    self.status_message = Some(format!("Redo: recreated '{}'", name));
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
                UndoAction::SortCleared { .. } => {
                    // Redo: clear the sort again
                    self.row_view.clear_sort();
                    self.filter_state.sort = None;
                    self.filter_state.invalidate_all_caches();
                    self.clamp_selection();
                    self.status_message = Some("Redo: clear sort".to_string());
                }
                UndoAction::ValidationSet { sheet_index, range, new_rule, .. } => {
                    // Redo: full replace pipeline - clear overlaps then set rule
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.validations.clear_range(&range);
                        sheet.validations.set(range.clone(), new_rule);
                    }
                    self.bump_cells_rev();
                    // Recompute invalid markers for affected range
                    self.revalidate_range(&range);
                    self.status_message = Some(format!("Redo: set validation ({} cells)", range.cell_count()));
                }
                UndoAction::ValidationCleared { sheet_index, range, .. } => {
                    // Redo: clear the validations again
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.validations.clear_range(&range);
                    }
                    self.bump_cells_rev();
                    // Clear invalid markers in the range (no validation = no invalid)
                    self.clear_invalid_markers_in_range(&range);
                    self.status_message = Some(format!("Redo: clear validation ({} cells)", range.cell_count()));
                }
                UndoAction::ValidationExcluded { sheet_index, range } => {
                    // Redo: re-add the exclusion
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.validations.exclude(range.clone());
                    }
                    self.bump_cells_rev();
                    // Clear invalid markers (excluded cells are not validated)
                    self.clear_invalid_markers_in_range(&range);
                    self.status_message = Some(format!("Redo: exclude from validation ({} cells)", range.cell_count()));
                }
                UndoAction::ValidationExclusionCleared { sheet_index, range, .. } => {
                    // Redo: clear exclusions again
                    if let Some(sheet) = self.workbook.sheet_mut(sheet_index) {
                        sheet.validations.clear_exclusions_in_range(&range);
                    }
                    self.bump_cells_rev();
                    // Recompute invalid markers (cells may now be validated)
                    self.revalidate_range(&range);
                    self.status_message = Some(format!("Redo: clear exclusions ({} cells)", range.cell_count()));
                }
                UndoAction::Rewind { .. } => {
                    // Rewind is audit-only - cannot be redone
                    self.status_message = Some("Cannot redo rewind".to_string());
                    return;  // Don't modify state
                }
            }
            self.is_modified = true;
            self.request_title_refresh(cx);
        }
    }

    /// Recompute invalid markers for a range after validation changes.
    /// Called after undo/redo to keep markers consistent with current validation state.
    fn revalidate_range(&mut self, range: &visigrid_engine::validation::CellRange) {
        use visigrid_engine::validation::ValidationResult;
        use visigrid_engine::workbook::Workbook;

        let sheet_index = self.sheet_index();

        // Clear existing markers in range first
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                self.invalid_cells.remove(&(row, col));
                self.validation_failures.retain(|&(r, c)| r != row || c != col);
            }
        }

        // Revalidate cells in range
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                let display_value = self.sheet().get_display(row, col);
                if display_value.is_empty() {
                    continue;
                }
                let result = self.workbook.validate_cell_input(sheet_index, row, col, &display_value);
                if let ValidationResult::Invalid { reason, .. } = result {
                    let failure_reason = Workbook::classify_failure_reason(&reason);
                    self.invalid_cells.insert((row, col), failure_reason);
                    if !self.validation_failures.contains(&(row, col)) {
                        self.validation_failures.push((row, col));
                    }
                }
            }
        }

        // Re-sort failures in row-major order
        self.validation_failures.sort_by_key(|&(r, c)| (r, c));
    }

    /// Clear invalid markers in a range (used when validation is removed).
    fn clear_invalid_markers_in_range(&mut self, range: &visigrid_engine::validation::CellRange) {
        for row in range.start_row..=range.end_row {
            for col in range.start_col..=range.end_col {
                self.invalid_cells.remove(&(row, col));
                self.validation_failures.retain(|&(r, c)| r != row || c != col);
            }
        }
    }
}
