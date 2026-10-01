//! Sort and filter operations
//!
//! Contains:
//! - Row View Layer (view space <-> data space conversion)
//! - Sort operations (sort by column, clear sort)
//! - AutoFilter (toggle, dropdown, apply filters)
//! - Filter helpers

use gpui::*;
use visigrid_engine::provenance::{MutationOp, SortKey};
use crate::app::Spreadsheet;

impl Spreadsheet {
    // ========================================================================
    // Row View Layer (view space <-> data space conversion)
    // ========================================================================
    // UI uses VIEW space (what user sees after sort/filter)
    // Storage uses DATA space (canonical row numbers)
    // Convert at boundaries only
    //
    // PREVIEW MODE: When previewing, we use the preview session's row_order
    // instead of the live row_view. This allows showing sorted state from history.

    /// Get the preview row order for the active sheet, if in preview mode
    /// Returns None if not previewing or no sort order recorded
    #[inline]
    pub fn preview_row_order(&self, cx: &App) -> Option<&[usize]> {
        self.preview_session().and_then(|session| {
            let sheet_idx = self.sheet_index(cx);
            session.view_state.per_sheet.get(sheet_idx)
                .and_then(|sv| sv.row_order.as_deref())
        })
    }

    /// Get the preview sort state for the active sheet
    /// Returns (column, is_ascending) if sort is active in preview
    #[inline]
    pub fn preview_sort_state(&self, cx: &App) -> Option<(usize, bool)> {
        self.preview_session().and_then(|session| {
            let sheet_idx = self.sheet_index(cx);
            session.view_state.per_sheet.get(sheet_idx)
                .and_then(|sv| sv.sort)
        })
    }

    /// Get the sort state to display (preview-aware)
    /// Returns (column, is_ascending) if the current view is sorted
    /// Uses preview sort state when previewing, live filter_state.sort otherwise
    #[inline]
    pub fn display_sort_state(&self, cx: &App) -> Option<(usize, bool)> {
        if self.is_previewing() {
            self.preview_sort_state(cx)
        } else {
            self.filter_state.sort.as_ref().map(|s| {
                (s.column, s.direction == visigrid_engine::filter::SortDirection::Ascending)
            })
        }
    }

    /// Convert view row to data row
    /// In preview mode, uses preview row order if available
    #[inline]
    pub fn view_to_data(&self, view_row: usize, cx: &App) -> usize {
        if let Some(row_order) = self.preview_row_order(cx) {
            // Preview mode with sorted state
            row_order.get(view_row).copied().unwrap_or(view_row)
        } else if self.is_previewing() {
            // Preview mode but sort invalidated - identity
            view_row
        } else {
            // Live mode
            self.row_view.view_to_data(view_row)
        }
    }

    /// Convert data row to view row (None if hidden by filter)
    /// In preview mode, filtering is not active (all rows visible)
    #[inline]
    pub fn data_to_view(&self, data_row: usize, cx: &App) -> Option<usize> {
        if let Some(row_order) = self.preview_row_order(cx) {
            // Preview mode with sorted state - find position
            row_order.iter().position(|&r| r == data_row)
        } else if self.is_previewing() {
            // Preview mode but sort invalidated - identity, always visible
            Some(data_row)
        } else {
            // Live mode
            self.row_view.data_to_view(data_row)
        }
    }

    /// Get visible row count
    /// In preview mode, all rows are considered visible (no filter state)
    #[inline]
    pub fn visible_row_count(&self) -> usize {
        if self.is_previewing() {
            // Preview doesn't track filters, use sheet's row count
            // (Heuristic: use same count as live row_view for consistency)
            self.row_view.row_count()
        } else {
            self.row_view.visible_count()
        }
    }

    /// Get the nth visible row (view_row, data_row) for rendering
    /// Returns None if index is out of bounds
    /// In preview mode, uses preview row order if available
    #[inline]
    pub fn nth_visible_row(&self, visible_index: usize, cx: &App) -> Option<(usize, usize)> {
        if let Some(row_order) = self.preview_row_order(cx) {
            // Preview mode with sorted state
            let data_row = row_order.get(visible_index).copied()?;
            // view_row == visible_index in sorted preview (no filter)
            Some((visible_index, data_row))
        } else if self.is_previewing() {
            // Preview mode but sort invalidated - identity mapping
            if visible_index < self.row_view.row_count() {
                Some((visible_index, visible_index))
            } else {
                None
            }
        } else {
            // Live mode
            let view_row = self.row_view.nth_visible(visible_index)?;
            let data_row = self.row_view.view_to_data(view_row);
            Some((view_row, data_row))
        }
    }

    /// View row indices that are visible after filtering
    /// (Not to be confused with visible_rows() which returns screen row count)
    #[inline]
    pub fn filtered_row_indices(&self) -> &[usize] {
        // Note: This method is used for live filtering state only
        // Preview mode doesn't use this - it uses preview_row_order() directly
        self.row_view.visible_rows()
    }

    /// Ensure row_view has enough capacity for current sheet
    pub fn ensure_row_view_capacity(&mut self) {
        // For now, use a large default. Later this can track actual data extent.
        let needed = 100000;
        if self.row_view.row_count() < needed {
            self.row_view.resize(needed);
        }
    }

    // ========================================================================
    // Sort Operations
    // ========================================================================

    /// Sort by the column containing the active cell
    ///
    /// This is the ONE command UI calls. Engine owns the permutation.
    pub fn sort_by_current_column(
        &mut self,
        direction: visigrid_engine::filter::SortDirection,
        cx: &mut Context<Self>,
    ) {
        use visigrid_engine::filter::{sort_by_column, SortState};

        // Block during preview mode
        if self.block_if_previewing(cx) { return; }

        // The table being sorted: the existing filter range, or the table
        // around the cursor (header = its first row, a title row above skipped).
        let range = match self.filter_state.filter_range {
            Some(r) => r,
            None => match self.table_range_at_cursor(cx) {
                Some(r) => r,
                None => {
                    self.status_message = Some("No data to sort".to_string());
                    cx.notify();
                    return;
                }
            },
        };
        // TODO(engine): enforce in engine sort API too once available (UI guard is not sufficient for headless).
        if self.block_if_merges_in_rows("sort", (range.0 + 1, range.2), cx) { return; }
        self.filter_state.filter_range = Some(range);

        let col = self.view_state.selected.1;

        // Create value_at closure (captures sheet for computed values)
        let sheet = self.sheet(cx);
        let value_at = |data_row: usize, c: usize| -> visigrid_engine::formula::eval::Value {
            sheet.get_computed_value(data_row, c)
        };

        // Call engine's sort function
        let (new_order, undo_item) = sort_by_column(
            &self.row_view,
            &self.filter_state,
            value_at,
            col,
            direction,
        );

        // Record undo for sort operation
        let previous_sort_state = undo_item.previous_sort_state.map(|s| {
            (s.column, s.direction == visigrid_engine::filter::SortDirection::Ascending)
        });
        let is_ascending = direction == visigrid_engine::filter::SortDirection::Ascending;

        // Build provenance
        let provenance = if let Some((start_row, start_col, end_row, end_col)) = self.filter_state.filter_range {
            Some(MutationOp::Sort {
                sheet: self.sheet(cx).id,
                range_start_row: start_row,
                range_start_col: start_col,
                range_end_row: end_row,
                range_end_col: end_col,
                keys: vec![SortKey { col, ascending: is_ascending }],
                has_header: false,  // TODO: detect header row
            }.to_provenance(&self.sheet(cx).name))
        } else {
            None
        };

        self.history.record_action_with_provenance(crate::history::UndoAction::SortApplied {
            sheet_index: self.sheet_index(cx),
            previous_row_order: undo_item.previous_row_order,
            previous_sort_state,
            new_row_order: new_order.clone(),
            new_sort_state: (col, is_ascending),
        }, provenance);

        // Apply the sort
        self.row_view.apply_sort(new_order);

        // Update filter_state.sort
        self.filter_state.sort = Some(SortState { column: col, direction });

        // Invalidate caches
        self.filter_state.invalidate_all_caches();

        self.is_modified = true;
        self.status_message = Some(format!(
            "Sorted by column {} {}",
            Self::col_to_letter(col),
            if direction == visigrid_engine::filter::SortDirection::Ascending { "A→Z" } else { "Z→A" }
        ));
        cx.notify();
    }

    /// The table around the cursor, as (header_row, first_col, last_row,
    /// last_col): the contiguous block of data the cursor is in. A top row with
    /// a single filled cell over a wider table is a title, not the header, and
    /// is skipped (a merged or centered report title sits there). None if
    /// there's no header with at least one row of data under it.
    pub(crate) fn table_range_at_cursor(&self, cx: &App) -> Option<(usize, usize, usize, usize)> {
        let (row, col) = self.view_state.selected;
        table_range_around(self.sheet(cx), row, col)
    }

    /// Toggle AutoFilter on/off for current selection
    pub fn toggle_auto_filter(&mut self, cx: &mut Context<Self>) {
        if self.filter_state.is_enabled() {
            // Disable: restore original order, clear filters
            self.row_view.clear_sort();
            self.row_view.clear_filter();
            self.filter_state.disable();
            self.status_message = Some("AutoFilter disabled".to_string());
        } else {
            // Enable on the table around the cursor (header = its first row).
            let Some(range) = self.table_range_at_cursor(cx) else {
                self.status_message = Some("No data for AutoFilter".to_string());
                cx.notify();
                return;
            };
            // TODO(engine): enforce in engine filter API too once available (UI guard is not sufficient for headless).
            if self.block_if_merges_in_rows("turn on AutoFilter", (range.0 + 1, range.2), cx) { return; }

            self.filter_state.filter_range = Some(range);
            self.status_message = Some(format!(
                "AutoFilter enabled: {}{}:{}{}",
                Self::col_to_letter(range.1),
                range.0 + 1,
                Self::col_to_letter(range.3),
                range.2 + 1
            ));
        }
        cx.notify();
    }

    /// Open the filter dropdown for a column
    pub fn open_filter_dropdown(&mut self, col: usize, cx: &mut Context<Self>) {
        if !self.filter_state.is_enabled() {
            return;
        }

        // Collect values for the column first (to avoid borrow conflicts)
        let values: Vec<(usize, visigrid_engine::formula::eval::Value)> = if let Some((data_start, _, data_end, _)) = self.filter_state.data_range() {
            let sheet = self.sheet(cx);
            (data_start..=data_end)
                .map(|data_row| (data_row, sheet.get_computed_value(data_row, col)))
                .collect()
        } else {
            Vec::new()
        };

        // Build unique values cache (max 500 unique values, sorted by frequency)
        self.filter_state.build_unique_values_from_vec(col, &values, 500);

        // Initialize checked items: all checked if no filter, or match current filter
        self.filter_checked_items.clear();
        if let Some(unique_vals) = self.filter_state.get_unique_values(col) {
            let col_filter = self.filter_state.column_filters.get(&col);
            for (idx, entry) in unique_vals.iter().enumerate() {
                // If no filter for this column, all items are checked
                // If filter exists, check if this value is in selected set
                let should_check = match col_filter {
                    None => true,
                    Some(cf) => match &cf.selected {
                        None => true, // No selection = all pass
                        Some(selected) => selected.contains(&entry.key),
                    },
                };
                if should_check {
                    self.filter_checked_items.insert(idx);
                }
            }
        }

        self.filter_dropdown_col = Some(col);
        self.filter_search_text.clear();
        cx.notify();
    }

    /// Close the filter dropdown without applying
    pub fn close_filter_dropdown(&mut self, cx: &mut Context<Self>) {
        self.filter_dropdown_col = None;
        self.filter_search_text.clear();
        self.filter_checked_items.clear();
        cx.notify();
    }

    /// Toggle a value in the filter dropdown
    pub fn toggle_filter_item(&mut self, idx: usize, cx: &mut Context<Self>) {
        if self.filter_checked_items.contains(&idx) {
            self.filter_checked_items.remove(&idx);
        } else {
            self.filter_checked_items.insert(idx);
        }
        cx.notify();
    }

    /// Select all items in filter dropdown
    pub fn filter_select_all(&mut self, cx: &mut Context<Self>) {
        let Some(col) = self.filter_dropdown_col else { return };
        if let Some(unique_vals) = self.filter_state.get_unique_values(col) {
            self.filter_checked_items.clear();
            for idx in 0..unique_vals.len() {
                self.filter_checked_items.insert(idx);
            }
        }
        cx.notify();
    }

    /// Clear all items in filter dropdown
    pub fn filter_clear_all(&mut self, cx: &mut Context<Self>) {
        self.filter_checked_items.clear();
        cx.notify();
    }

    /// Apply the current filter dropdown selection
    pub fn apply_filter_dropdown(&mut self, cx: &mut Context<Self>) {
        let Some(col) = self.filter_dropdown_col else { return };
        let Some(unique_vals) = self.filter_state.get_unique_values(col) else {
            self.close_filter_dropdown(cx);
            return;
        };

        // Build selected set from checked items
        let all_checked = self.filter_checked_items.len() == unique_vals.len();

        if all_checked {
            // All checked = no filter (remove filter for this column)
            self.filter_state.clear_column_filter(col);
        } else {
            // Build HashSet of selected normalized keys
            let selected: std::collections::HashSet<_> = self
                .filter_checked_items
                .iter()
                .filter_map(|&idx| unique_vals.get(idx).map(|e| e.key.clone()))
                .collect();

            self.filter_state.column_filters.insert(
                col,
                visigrid_engine::filter::ColumnFilter {
                    selected: Some(selected),
                    text_filter: None,
                },
            );
        }

        // Apply filters to row_view
        self.apply_all_filters(cx);

        self.filter_dropdown_col = None;
        self.filter_search_text.clear();
        self.filter_checked_items.clear();
        self.is_modified = true;
        cx.notify();
    }

    /// Apply all column filters to update visible_mask
    fn apply_all_filters(&mut self, cx: &App) {
        let Some((data_start, min_col, data_end, max_col)) = self.filter_state.data_range() else {
            // No filter range - all visible
            self.row_view.clear_filter();
            return;
        };

        // Build visible_mask for all data rows
        let row_count = self.row_view.row_count();
        let mut visible_mask = vec![true; row_count];

        // Header row always visible
        if let Some(header) = self.filter_state.header_row() {
            visible_mask[header] = true;
        }

        // Check each data row against all column filters
        for data_row in data_start..=data_end {
            if data_row >= row_count {
                break;
            }

            let mut passes = true;
            for col in min_col..=max_col {
                if let Some(col_filter) = self.filter_state.column_filters.get(&col) {
                    if col_filter.is_active() {
                        let value = self.sheet(cx).get_computed_value(data_row, col);
                        let filter_key = visigrid_engine::filter::FilterKey::from_value(&value);
                        if !col_filter.passes(&filter_key) {
                            passes = false;
                            break;
                        }
                    }
                }
            }
            visible_mask[data_row] = passes;
        }

        self.row_view.apply_filter(visible_mask);
    }

    /// Check if a column has an active filter
    pub fn column_has_filter(&self, col: usize) -> bool {
        self.filter_state
            .column_filters
            .get(&col)
            .map_or(false, |f| f.is_active())
    }

    /// Clear sort (restore original data order)
    pub fn clear_sort(&mut self, cx: &mut Context<Self>) {
        // Only record undo if there's actually a sort to clear
        if let Some(sort_state) = &self.filter_state.sort {
            // Capture previous state for undo
            let previous_row_order = self.row_view.row_order().to_vec();
            let previous_sort_state = (
                sort_state.column,
                sort_state.direction == visigrid_engine::filter::SortDirection::Ascending,
            );

            // Record undo action (no provenance for clear)
            self.history.record_action_with_provenance(crate::history::UndoAction::SortCleared {
                sheet_index: self.sheet_index(cx),
                previous_row_order,
                previous_sort_state,
            }, None);
        }

        self.row_view.clear_sort();
        self.filter_state.sort = None;
        self.is_modified = true;
        self.status_message = Some("Sort cleared".to_string());
        cx.notify();
    }
}

/// The table around (row, col) as (header_row, first_col, last_row, last_col).
/// A top row with a single filled cell over a wider table is a title, not the
/// header, and is skipped. None without a header and at least one data row.
pub(crate) fn table_range_around(
    sheet: &visigrid_engine::sheet::Sheet,
    row: usize,
    col: usize,
) -> Option<(usize, usize, usize, usize)> {
    let (mut r0, c0, r1, c1) = crate::ai::find_current_region(sheet, row, col);
    let filled = |r: usize| (c0..=c1).filter(|&c| !sheet.get_display(r, c).is_empty()).count();
    while c1 > c0 && r1 > r0 + 1 && filled(r0) <= 1 {
        r0 += 1;
    }
    (r1 > r0).then_some((r0, c0, r1, c1))
}

#[cfg(test)]
mod table_range_tests {
    use super::table_range_around;
    use visigrid_engine::sheet::{Sheet, SheetId};

    fn sheet(rows: &[&[&str]]) -> Sheet {
        let mut s = Sheet::new(SheetId(1), 100, 10);
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                if !v.is_empty() {
                    // A standalone test sheet with no workbook or dependents,
                    // so the untracked write is fine (the tests::no_untracked_
                    // cell_mutations scan looks for the method-call form).
                    Sheet::set_value(&mut s, r, c, v);
                }
            }
        }
        s
    }

    #[test]
    fn a_title_row_above_the_header_is_skipped() {
        let s = sheet(&[
            &["Q3 Sales Report", "", ""],
            &["Region", "Rep", "Amount"],
            &["West", "Ann", "10"],
            &["East", "Bo", "5"],
        ]);
        assert_eq!(table_range_around(&s, 2, 1), Some((1, 0, 3, 2)));
        // From the title cell too.
        assert_eq!(table_range_around(&s, 0, 0), Some((1, 0, 3, 2)));
    }

    #[test]
    fn a_table_separated_from_its_title_and_a_plain_table() {
        let s = sheet(&[&["Report", "", ""], &["", "", ""], &["Region", "Rep", "Amount"], &["West", "Ann", "10"]]);
        assert_eq!(table_range_around(&s, 3, 0), Some((2, 0, 3, 2)));
        let plain = sheet(&[&["Region", "Amount"], &["West", "10"]]);
        assert_eq!(table_range_around(&plain, 1, 1), Some((0, 0, 1, 1)));
        // A single column has no title to skip.
        let one = sheet(&[&["Names"], &["Ann"], &["Bo"]]);
        assert_eq!(table_range_around(&one, 1, 0), Some((0, 0, 2, 0)));
        // A header with no data under it isn't sortable.
        let header_only = sheet(&[&["Region", "Amount"]]);
        assert_eq!(table_range_around(&header_only, 0, 0), None);
    }
}

